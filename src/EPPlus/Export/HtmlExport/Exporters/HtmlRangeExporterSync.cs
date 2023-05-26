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
using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class HtmlRangeExporterSync : HtmlRangeExporterSyncBase
    {
        internal HtmlRangeExporterSync
            (HtmlRangeExportSettings settings, ExcelRangeBase range) : base(settings, range)
        {
            this._settings = settings;
        }

        internal HtmlRangeExporterSync(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges) : base(settings, ranges)
        {
            this._settings = settings;
        }

        private readonly HtmlRangeExportSettings _settings;

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString()
        {
            using MemoryStream? ms = RecyclableMemory.GetStream();
            this.RenderHtml(ms, 0);
            ms.Position = 0;
            using StreamReader? sr = new StreamReader(ms);
            return sr.ReadToEnd();
        }
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString(int rangeIndex)
        {
            this.ValidateRangeIndex(rangeIndex);
            using MemoryStream? ms = RecyclableMemory.GetStream();
            this.RenderHtml(ms, rangeIndex);
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
        public string GetHtmlString(int rangeIndex, ExcelHtmlOverrideExportSettings settings)
        {
            this.ValidateRangeIndex(rangeIndex);
            using MemoryStream? ms = RecyclableMemory.GetStream();
            this.RenderHtml(ms, rangeIndex, settings);
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
        public string GetHtmlString(int rangeIndex, Action<ExcelHtmlOverrideExportSettings> config)
        {
            ExcelHtmlOverrideExportSettings? settings = new ExcelHtmlOverrideExportSettings();
            config.Invoke(settings);
            return this.GetHtmlString(rangeIndex, settings);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <returns>A html table</returns>
        public void RenderHtml(Stream stream)
        {
            this.RenderHtml(stream, 0);
        }
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <param name="rangeIndex">The index of the range to output.</param>
        /// <param name="overrideSettings">Settings for this specific range index</param>
        /// <returns>A html table</returns>
        public void RenderHtml(Stream stream, int rangeIndex, ExcelHtmlOverrideExportSettings overrideSettings = null)
        {
            this.ValidateRangeIndex(rangeIndex);

            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }

            this._mergedCells.Clear();
            ExcelRangeBase? range = this._ranges[rangeIndex];
            this.GetDataTypes(range, this._settings);

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
            writer.RenderBeginTag(HtmlElements.Table);

            writer.ApplyFormatIncreaseIndent(this.Settings.Minify);
            this.LoadVisibleColumns(range);
            if (this.Settings.SetColumnWidth || this.Settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
            {
                this.SetColumnGroup(writer, range, this.Settings, this.IsMultiSheet);
            }

            if (this._settings.HeaderRows > 0 || this._settings.Headers.Count > 0)
            {
                this.RenderHeaderRow(range, writer, table, accessibilitySettings, headerRows, headers);
            }
            // table rows
            this.RenderTableRows(range, writer, table, accessibilitySettings);

            writer.ApplyFormatDecreaseIndent(this.Settings.Minify);
            // end tag table
            writer.RenderEndTag();
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <param name="rangeIndex">Index of the range to export</param>
        /// <param name="config">Override some of the settings for this html exclusively</param>
        /// <returns></returns>
        public void RenderHtml(Stream stream, int rangeIndex, Action<ExcelHtmlOverrideExportSettings> config)
        {
            ExcelHtmlOverrideExportSettings? settings = new ExcelHtmlOverrideExportSettings();
            config.Invoke(settings);
            this.RenderHtml(stream, rangeIndex, settings);
        }

        /// <summary>
        /// The ranges used in the export.
        /// </summary>
        public EPPlusReadOnlyList<ExcelRangeBase> Ranges
        {
            get
            {
                return this._ranges;
            }
        }

        /// <summary>
        /// Renders both the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public string GetSinglePage(string htmlDocument = "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}\r\n</body>\r\n</html>")
        {
            if (this.Settings.Minify)
            {
                htmlDocument = htmlDocument.Replace("\r\n", "");
            }

            string? html = this.GetHtmlString();
            CssRangeExporterSync? exporter = HtmlExporterFactory.CreateCssExporterSync(this._settings, this._ranges, this._styleCache);
            string? css = exporter.GetCssString();
            return string.Format(htmlDocument, html, css);
        }

        private void RenderTableRows(ExcelRangeBase range, EpplusHtmlWriter writer, ExcelTable table, AccessibilitySettings accessibilitySettings)
        {
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TbodyRole))
            {
                writer.AddAttribute("role", accessibilitySettings.TableSettings.TbodyRole);
            }
            writer.RenderBeginTag(HtmlElements.Tbody);
            writer.ApplyFormatIncreaseIndent(this.Settings.Minify);
            int row = range._fromRow + this._settings.HeaderRows;
            int endRow = range._toRow;
            ExcelWorksheet? ws = range.Worksheet;
            HtmlImage image = null;
            bool hasFooter = table != null && table.ShowTotal;
            while (row <= endRow)
            {
                if (HandleHiddenRow(writer, range.Worksheet, this.Settings, ref row))
                {
                    continue; //The row is hidden and should not be included.
                }

                if (hasFooter && row == endRow)
                {
                    writer.RenderBeginTag(HtmlElements.TFoot);
                }

                if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "row");
                    writer.AddAttribute("scope", "row");
                }

                if (this.Settings.SetRowHeight)
                {
                    AddRowHeightStyle(writer, range, row, this.Settings.StyleClassPrefix, this.IsMultiSheet);
                }

                writer.RenderBeginTag(HtmlElements.TableRow);
                writer.ApplyFormatIncreaseIndent(this.Settings.Minify);
                foreach (int col in this._columns)
                {
                    if (this.InMergeCellSpan(row, col))
                    {
                        continue;
                    }

                    int colIx = col - range._fromCol;
                    ExcelRange? cell = ws.Cells[row, col];
                    object? cv = cell.Value;
                    string? dataType = HtmlRawDataProvider.GetHtmlDataTypeFromValue(cell.Value);

                    this.SetColRowSpan(range, writer, cell);

                    if (this.Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = this.GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                    }

                    if (cell.Hyperlink == null)
                    {
                        CellDataWriter.Write(cell, dataType, writer, this.Settings, accessibilitySettings, false, image);
                    }
                    else
                    {
                        string? imageCellClassName = GetImageCellClassName(image, this.Settings);
                        writer.SetClassAttributeFromStyle(cell, false, this.Settings, imageCellClassName);
                        writer.RenderBeginTag(HtmlElements.TableData);
                        AddImage(writer, this.Settings, image, cell.Value);
                        RenderHyperlink(writer, cell, this.Settings);
                        writer.RenderEndTag();
                        writer.ApplyFormat(this.Settings.Minify);
                    }
                }

                // end tag tr
                writer.Indent--;
                writer.RenderEndTag();
                writer.ApplyFormat(this.Settings.Minify);
                if (hasFooter && row == endRow)
                {
                    writer.RenderEndTag();
                }
                row++;
            }

            writer.ApplyFormatDecreaseIndent(this.Settings.Minify);
            // end tag tbody
            writer.RenderEndTag();
        }
        private void RenderHeaderRow(ExcelRangeBase range, EpplusHtmlWriter writer, ExcelTable table, AccessibilitySettings accessibilitySettings, int headerRows, List<string> headers)
        {
            if (table != null && table.ShowHeader == false)
            {
                return;
            }

            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TheadRole))
            {
                writer.AddAttribute("role", this.Settings.Accessibility.TableSettings.TheadRole);
            }
            writer.RenderBeginTag(HtmlElements.Thead);
            writer.ApplyFormatIncreaseIndent(this.Settings.Minify);
            if (table == null)
            {
                headerRows = this._settings.HeaderRows == 0 ? 1 : this._settings.HeaderRows;
            }
            else
            {
                headerRows = table.ShowHeader ? 1 : 0;
            }

            HtmlImage image = null;
            for (int i = 0; i < headerRows; i++)
            {
                if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "row");
                }
                int row = range._fromRow + i;
                if (this.Settings.SetRowHeight)
                {
                    AddRowHeightStyle(writer, range, row, this.Settings.StyleClassPrefix, this.IsMultiSheet);
                }

                writer.RenderBeginTag(HtmlElements.TableRow);
                writer.ApplyFormatIncreaseIndent(this.Settings.Minify);
                foreach (int col in this._columns)
                {
                    if (this.InMergeCellSpan(row, col))
                    {
                        continue;
                    }

                    ExcelRange? cell = range.Worksheet.Cells[row, col];
                    if (this.Settings.RenderDataTypes)
                    {
                        writer.AddAttribute("data-datatype", this._dataTypes[col - range._fromCol]);
                    }

                    this.SetColRowSpan(range, writer, cell);
                    if (this.Settings.IncludeCssClassNames)
                    {
                        string? imageCellClassName = GetImageCellClassName(image, this.Settings);
                        writer.SetClassAttributeFromStyle(cell, true, this.Settings, imageCellClassName);
                    }
                    if (this.Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = this.GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                    }
                    writer.RenderBeginTag(HtmlElements.TableHeader);
                    AddImage(writer, this.Settings, image, cell.Value);

                    if (headerRows > 0 || table != null)
                    {
                        if (cell.Hyperlink == null)
                        {
                            writer.Write(GetCellText(cell, this.Settings));
                        }
                        else
                        {
                            RenderHyperlink(writer, cell, this.Settings);
                        }
                    }
                    else if (headers.Count < col)
                    {
                        writer.Write(headers[col]);
                    }

                    writer.RenderEndTag();
                    writer.ApplyFormat(this.Settings.Minify);
                }
                writer.Indent--;
                writer.RenderEndTag();
            }
            writer.ApplyFormatDecreaseIndent(this.Settings.Minify);
            writer.RenderEndTag();
            writer.ApplyFormat(this.Settings.Minify);
        }
    }
}
