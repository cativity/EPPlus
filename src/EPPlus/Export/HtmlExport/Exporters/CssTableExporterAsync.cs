﻿/*************************************************************************************************
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

using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
#if !NET35 && !NET40
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class CssTableExporterAsync : CssRangeExporterBase
    {
        public CssTableExporterAsync(HtmlTableExportSettings settings, ExcelTable table)
            : base(settings, table.Range)
        {
            this._table = table;
            this._tableSettings = settings;
        }

        private readonly ExcelTable _table;
        private readonly HtmlTableExportSettings _tableSettings;

        /// <summary>
        /// Exports the css part of an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public async Task<string> GetCssStringAsync()
        {
            using MemoryStream? ms = RecyclableMemory.GetStream();
            await this.RenderCssAsync(ms);
            ms.Position = 0;
            using StreamReader? sr = new StreamReader(ms);

            return await sr.ReadToEndAsync();
        }

        /// <summary>
        /// Exports the css part of an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public async Task RenderCssAsync(Stream stream)
        {
            if ((this._table.TableStyle == TableStyles.None || this._tableSettings.Css.IncludeTableStyles == false)
                && this._tableSettings.Css.IncludeCellStyles == false)
            {
                return;
            }

            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }

            if (this._dataTypes.Count == 0)
            {
                this.GetDataTypes(this._table.Address, this._table);
            }

            StreamWriter? sw = new StreamWriter(stream);
            List<ExcelRangeBase>? ranges = new List<ExcelRangeBase>() { this._table.Range };

            EpplusCssWriter? cellCssWriter = new EpplusCssWriter(sw,
                                                                 ranges,
                                                                 this._tableSettings,
                                                                 this._tableSettings.Css,
                                                                 this._tableSettings.Css.Exclude.CellStyle,
                                                                 this._styleCache);

            await cellCssWriter.RenderAdditionalAndFontCssAsync(TableClass);

            if (this._tableSettings.Css.IncludeTableStyles)
            {
                await RenderTableCssAsync(sw, this._table, this._tableSettings, this._styleCache, this._dataTypes);
            }

            if (this._tableSettings.Css.IncludeCellStyles)
            {
                await this.RenderCellCssAsync(sw);
            }

            if (this.Settings.Pictures.Include == ePictureInclude.Include)
            {
                this.LoadRangeImages(ranges);

                foreach (HtmlImage? p in this._rangePictures)
                {
                    await cellCssWriter.AddPictureToCssAsync(p);
                }
            }

            await cellCssWriter.FlushStreamAsync();
        }

        private async Task RenderCellCssAsync(StreamWriter sw)
        {
            List<ExcelRangeBase>? ranges = new List<ExcelRangeBase>() { this._table.Range };

            EpplusCssWriter? styleWriter = new EpplusCssWriter(sw,
                                                               ranges,
                                                               this._tableSettings,
                                                               this._tableSettings.Css,
                                                               this._tableSettings.Css.Exclude.CellStyle,
                                                               this._styleCache);

            ExcelRangeBase? r = this._table.Range;
            ExcelStyles? styles = r.Worksheet.Workbook.Styles;
            CellStoreEnumerator<ExcelValue>? ce = new CellStoreEnumerator<ExcelValue>(r.Worksheet._values, r._fromRow, r._fromCol, r._toRow, r._toCol);

            while (ce.Next())
            {
                if (ce.Value._styleId > 0 && ce.Value._styleId < styles.CellXfs.Count)
                {
                    await styleWriter.AddToCssAsync(styles, ce.Value._styleId, this.Settings.StyleClassPrefix, this.Settings.CellStyleClassName);
                }
            }

            await styleWriter.FlushStreamAsync();
        }

        internal static async Task RenderTableCssAsync(StreamWriter sw,
                                                       ExcelTable table,
                                                       HtmlTableExportSettings settings,
                                                       Dictionary<string, int> styleCache,
                                                       List<string> datatypes)
        {
            EpplusTableCssWriter? styleWriter = new EpplusTableCssWriter(sw, table, settings, styleCache);

            if (settings.Minify == false)
            {
                await styleWriter.WriteLineAsync();
            }

            ExcelTableNamedStyle tblStyle;

            if (table.TableStyle == TableStyles.Custom)
            {
                tblStyle = table.WorkSheet.Workbook.Styles.TableStyles[table.StyleName].As.TableStyle;
            }
            else
            {
                XmlElement? tmpNode = table.WorkSheet.Workbook.StylesXml.CreateElement("c:tableStyle");
                tblStyle = new ExcelTableNamedStyle(table.WorkSheet.Workbook.Styles.NameSpaceManager, tmpNode, table.WorkSheet.Workbook.Styles);
                tblStyle.SetFromTemplate(table.TableStyle);
            }

            string? tableClass =
                $"{TableClass}.{HtmlExportTableUtil.TableStyleClassPrefix}{HtmlExportTableUtil.GetClassName(tblStyle.Name, "EmptyClassName").ToLower()}";

            await styleWriter.AddHyperlinkCssAsync($"{tableClass}", tblStyle.WholeTable);
            await styleWriter.AddAlignmentToCssAsync($"{tableClass}", datatypes);

            await styleWriter.AddToCssAsync($"{tableClass}", tblStyle.WholeTable, "");
            await styleWriter.AddToCssBorderVHAsync($"{tableClass}", tblStyle.WholeTable, "");

            //Header
            await styleWriter.AddToCssAsync($"{tableClass}", tblStyle.HeaderRow, " thead");
            await styleWriter.AddToCssBorderVHAsync($"{tableClass}", tblStyle.HeaderRow, "");

            await styleWriter.AddToCssAsync($"{tableClass}", tblStyle.LastTotalCell, $" thead tr th:last-child)");
            await styleWriter.AddToCssAsync($"{tableClass}", tblStyle.FirstHeaderCell, " thead tr th:first-child");

            //Total
            await styleWriter.AddToCssAsync($"{tableClass}", tblStyle.TotalRow, " tfoot");
            await styleWriter.AddToCssBorderVHAsync($"{tableClass}", tblStyle.TotalRow, "");
            await styleWriter.AddToCssAsync($"{tableClass}", tblStyle.LastTotalCell, $" tfoot tr td:last-child)");
            await styleWriter.AddToCssAsync($"{tableClass}", tblStyle.FirstTotalCell, " tfoot tr td:first-child");

            //Columns stripes
            string? tableClassCS = $"{tableClass}-column-stripes";
            await styleWriter.AddToCssAsync($"{tableClassCS}", tblStyle.FirstColumnStripe, $" tbody tr td:nth-child(odd)");
            await styleWriter.AddToCssAsync($"{tableClassCS}", tblStyle.SecondColumnStripe, $" tbody tr td:nth-child(even)");

            //Row stripes
            string? tableClassRS = $"{tableClass}-row-stripes";
            await styleWriter.AddToCssAsync($"{tableClassRS}", tblStyle.FirstRowStripe, " tbody tr:nth-child(odd)");
            await styleWriter.AddToCssAsync($"{tableClassRS}", tblStyle.SecondRowStripe, " tbody tr:nth-child(even)");

            //Last column
            string? tableClassLC = $"{tableClass}-last-column";
            await styleWriter.AddToCssAsync($"{tableClassLC}", tblStyle.LastColumn, $" tbody tr td:last-child");

            //First column
            string? tableClassFC = $"{tableClass}-first-column";
            await styleWriter.AddToCssAsync($"{tableClassFC}", tblStyle.FirstColumn, " tbody tr td:first-child");

            await styleWriter.FlushStreamAsync();
        }
    }
}
#endif