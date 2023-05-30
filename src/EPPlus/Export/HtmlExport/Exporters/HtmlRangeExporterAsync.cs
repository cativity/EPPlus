/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  6/4/2022         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/

using OfficeOpenXml.Core;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml.Export.HtmlExport.Accessibility;
#if !NET35 && !NET40
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class HtmlRangeExporterAsync : HtmlRangeExporterAsyncBase
    {
        internal HtmlRangeExporterAsync(HtmlRangeExportSettings settings, ExcelRangeBase range)
            : base(settings, range) =>
            this._settings = settings;

        internal HtmlRangeExporterAsync(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges)
            : base(settings, ranges) =>
            this._settings = settings;

        private readonly HtmlRangeExportSettings _settings;

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public async Task<string> GetHtmlStringAsync()
        {
            using MemoryStream? ms = RecyclableMemory.GetStream();
            await this.RenderHtmlAsync(ms, 0);
            ms.Position = 0;
            using StreamReader? sr = new StreamReader(ms);

            return sr.ReadToEnd();
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="settings">Override some of the settings for this html exclusively</param>
        /// <returns>A html table</returns>
        public async Task<string> GetHtmlStringAsync(int rangeIndex, ExcelHtmlOverrideExportSettings settings = null)
        {
            using MemoryStream? ms = RecyclableMemory.GetStream();
            await this.RenderHtmlAsync(ms, rangeIndex, settings);
            ms.Position = 0;
            using StreamReader? sr = new StreamReader(ms);

            return sr.ReadToEnd();
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="config">Override some of the settings for this html exclusively</param>
        /// <returns></returns>
        public async Task<string> GetHtmlStringAsync(int rangeIndex, Action<ExcelHtmlOverrideExportSettings> config)
        {
            ExcelHtmlOverrideExportSettings? settings = new ExcelHtmlOverrideExportSettings();
            config.Invoke(settings);

            return await this.GetHtmlStringAsync(rangeIndex, settings);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <returns>A html table</returns>
        public async Task RenderHtmlAsync(Stream stream) => await this.RenderHtmlAsync(stream, 0);

        /// <summary>
        /// Exports the html part of the html export, without the styles.
        /// </summary>
        /// <param name="stream">The stream to write to.</param>
        /// <param name="rangeIndex">The index of the range to output.</param>
        /// <param name="overrideSettings">Settings for this specific range index</param>
        /// <exception cref="IOException"></exception>
        public async Task RenderHtmlAsync(Stream stream, int rangeIndex, ExcelHtmlOverrideExportSettings overrideSettings = null)
        {
            this.ValidateRangeIndex(rangeIndex);

            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }

            this._mergedCells.Clear();
            ExcelRangeBase? range = this._ranges[rangeIndex];
            this.GetDataTypes(this._ranges[rangeIndex], this._settings);

            ExcelTable table = null;

            if (this.Settings.TableStyle != eHtmlRangeTableInclude.Exclude)
            {
                table = range.GetTable();
            }

            EpplusHtmlWriter? writer = new EpplusHtmlWriter(stream, this.Settings.Encoding, this._styleCache);
            string? tableId = this.GetTableId(rangeIndex, overrideSettings);
            List<string>? additionalClassNames = this.GetAdditionalClassNames(overrideSettings);
            AccessibilitySettings? accessibilitySettings = this.GetAccessibilitySettings(overrideSettings);
            int headerRows = overrideSettings != null ? overrideSettings.HeaderRows : this._settings.HeaderRows;
            List<string>? headers = overrideSettings != null ? overrideSettings.Headers : this._settings.Headers;
            AddClassesAttributes(writer, table, tableId, additionalClassNames);
            AddTableAccessibilityAttributes(accessibilitySettings, writer);
            await writer.RenderBeginTagAsync(HtmlElements.Table);

            await writer.ApplyFormatIncreaseIndentAsync(this.Settings.Minify);
            this.LoadVisibleColumns(range);

            if (this.Settings.SetColumnWidth || this.Settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
            {
                await this.SetColumnGroupAsync(writer, range, this.Settings, this.IsMultiSheet);
            }

            if (headerRows > 0 || headers.Count > 0)
            {
                await this.RenderHeaderRowAsync(range, writer, table, accessibilitySettings, headerRows, headers);
            }

            // table rows
            await this.RenderTableRowsAsync(range, writer, table, accessibilitySettings, this._settings.HeaderRows);

            await writer.ApplyFormatDecreaseIndentAsync(this.Settings.Minify);

            // end tag table
            await writer.RenderEndTagAsync();
        }

        /// <summary>
        /// Exports the html part of the html export, without the styles.
        /// </summary>
        /// <param name="stream">The stream to write to.</param>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="config">Override some of the settings for this html exclusively</param>
        /// <returns></returns>
        public async Task RenderHtmlAsync(Stream stream, int rangeIndex, Action<ExcelHtmlOverrideExportSettings> config)
        {
            ExcelHtmlOverrideExportSettings? settings = new ExcelHtmlOverrideExportSettings();
            config.Invoke(settings);
            await this.RenderHtmlAsync(stream, rangeIndex, settings);
        }

        /// <summary>
        /// Renders the first range of the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public async Task<string> GetSinglePageAsync(string htmlDocument =
                                                         "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}\r\n</body>\r\n</html>")
        {
            if (this.Settings.Minify)
            {
                htmlDocument = htmlDocument.Replace("\r\n", "");
            }

            string? html = await this.GetHtmlStringAsync();
            CssRangeExporterAsync? cssExporter = HtmlExporterFactory.CreateCssExporterAsync(this._settings, this._ranges, this._styleCache);
            string? css = await cssExporter.GetCssStringAsync();

            return string.Format(htmlDocument, html, css);
        }
    }
}
#endif