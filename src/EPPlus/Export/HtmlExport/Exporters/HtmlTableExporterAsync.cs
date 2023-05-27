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

using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml.Sorting;
#if !NET35 && !NET40
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class HtmlTableExporterAsync : HtmlRangeExporterAsyncBase
    {
        public HtmlTableExporterAsync(HtmlTableExportSettings settings, ExcelTable table)
            : base(settings, table.Range)
        {
            Require.Argument(table).IsNotNull("table");
            this._table = table;
            this._settings = settings;
            this.LoadRangeImages(new List<ExcelRangeBase>() { table.Range });
        }

        private readonly ExcelTable _table;
        private HtmlTableExportSettings _settings;

        private void LoadVisibleColumns()
        {
            this._columns = new List<int>();
            ExcelRangeBase? r = this._table.Range;

            for (int col = r._fromCol; col <= r._toCol; col++)
            {
                ExcelColumn? c = this._table.WorkSheet.GetColumn(col);

                if (c == null || (c.Hidden == false && c.Width > 0))
                {
                    this._columns.Add(col);
                }
            }
        }

        private async Task RenderTableRowsAsync(EpplusHtmlWriter writer, AccessibilitySettings accessibilitySettings)
        {
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TbodyRole))
            {
                writer.AddAttribute("role", accessibilitySettings.TableSettings.TbodyRole);
            }

            await writer.RenderBeginTagAsync(HtmlElements.Tbody);
            await writer.ApplyFormatIncreaseIndentAsync(this.Settings.Minify);
            int row = this._table.ShowHeader ? this._table.Address._fromRow + 1 : this._table.Address._fromRow;
            int endRow = this._table.ShowTotal ? this._table.Address._toRow - 1 : this._table.Address._toRow;
            HtmlImage image = null;

            while (row <= endRow)
            {
                if (HandleHiddenRow(writer, this._table.WorkSheet, this.Settings, ref row))
                {
                    continue; //The row is hidden and should not be included.
                }

                if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "row");

                    if (!this._table.ShowFirstColumn && !this._table.ShowLastColumn)
                    {
                        writer.AddAttribute("scope", "row");
                    }
                }

                if (this.Settings.SetRowHeight)
                {
                    AddRowHeightStyle(writer, this._table.Range, row, this.Settings.StyleClassPrefix, false);
                }

                await writer.RenderBeginTagAsync(HtmlElements.TableRow);
                await writer.ApplyFormatIncreaseIndentAsync(this.Settings.Minify);

                foreach (int col in this._columns)
                {
                    int colIx = col - this._table.Address._fromCol;
                    string? dataType = this._dataTypes[colIx];
                    ExcelRange? cell = this._table.WorkSheet.Cells[row, col];

                    if (this.Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = this.GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                    }

                    if (cell.Hyperlink == null)
                    {
                        bool addRowScope = (this._table.ShowFirstColumn && col == this._table.Address._fromCol)
                                           || (this._table.ShowLastColumn && col == this._table.Address._toCol);

                        await CellDataWriter.WriteAsync(cell, dataType, writer, this.Settings, accessibilitySettings, addRowScope, image);
                    }
                    else
                    {
                        await writer.RenderBeginTagAsync(HtmlElements.TableData);
                        string? imageCellClassName = GetImageCellClassName(image, this.Settings);
                        writer.SetClassAttributeFromStyle(cell, false, this.Settings, imageCellClassName);
                        await RenderHyperlinkAsync(writer, cell, this.Settings);
                        await writer.RenderEndTagAsync();
                        await writer.ApplyFormatAsync(this.Settings.Minify);
                    }
                }

                // end tag tr
                writer.Indent--;
                await writer.RenderEndTagAsync();
                await writer.ApplyFormatAsync(this.Settings.Minify);
                row++;
            }

            await writer.ApplyFormatDecreaseIndentAsync(this.Settings.Minify);

            // end tag tbody
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatAsync(this.Settings.Minify);
        }

        private async Task RenderHeaderRowAsync(EpplusHtmlWriter writer, AccessibilitySettings accessibilitySettings)
        {
            // table header row
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TheadRole))
            {
                writer.AddAttribute("role", accessibilitySettings.TableSettings.TheadRole);
            }

            await writer.RenderBeginTagAsync(HtmlElements.Thead);
            await writer.ApplyFormatIncreaseIndentAsync(this.Settings.Minify);

            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "row");
            }

            ExcelAddressBase? adr = this._table.Address;
            int row = adr._fromRow;

            if (this.Settings.SetRowHeight)
            {
                AddRowHeightStyle(writer, this._table.Range, row, this.Settings.StyleClassPrefix, false);
            }

            await writer.RenderBeginTagAsync(HtmlElements.TableRow);
            await writer.ApplyFormatIncreaseIndentAsync(this.Settings.Minify);
            HtmlImage image = null;

            foreach (int col in this._columns)
            {
                ExcelRange? cell = this._table.WorkSheet.Cells[row, col];

                if (this.Settings.RenderDataTypes)
                {
                    writer.AddAttribute("data-datatype", this._dataTypes[col - adr._fromCol]);
                }

                string? imageCellClassName = image == null ? "" : this.Settings.StyleClassPrefix + "image-cell";
                writer.SetClassAttributeFromStyle(cell, true, this.Settings, imageCellClassName);

                if (accessibilitySettings.TableSettings.AddAccessibilityAttributes
                    && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TableHeaderCellRole))
                {
                    writer.AddAttribute("role", accessibilitySettings.TableSettings.TableHeaderCellRole);

                    if (!this._table.ShowFirstColumn && !this._table.ShowLastColumn)
                    {
                        writer.AddAttribute("scope", "col");
                    }

                    if (this._table.SortState != null && !this._table.SortState.ColumnSort && this._table.SortState.SortConditions.Any())
                    {
                        SortCondition? firstCondition = this._table.SortState.SortConditions.First();

                        if (firstCondition != null && !string.IsNullOrEmpty(firstCondition.Ref))
                        {
                            ExcelAddress? addr = new ExcelAddress(firstCondition.Ref);
                            int sortedCol = addr._fromCol;

                            if (col == sortedCol)
                            {
                                writer.AddAttribute("aria-sort", firstCondition.Descending ? "descending" : "ascending");
                            }
                        }
                    }
                }

                await writer.RenderBeginTagAsync(HtmlElements.TableHeader);

                if (this.Settings.Pictures.Include == ePictureInclude.Include)
                {
                    image = this.GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
                }

                await AddImageAsync(writer, this.Settings, image, cell.Value);

                if (cell.Hyperlink == null)
                {
                    await writer.WriteAsync(GetCellText(cell, this.Settings));
                }
                else
                {
                    await RenderHyperlinkAsync(writer, cell, this.Settings);
                }

                await writer.RenderEndTagAsync();
                await writer.ApplyFormatAsync(this.Settings.Minify);
            }

            writer.Indent--;
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatDecreaseIndentAsync(this.Settings.Minify);
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatAsync(this.Settings.Minify);
        }

        private async Task RenderTotalRowAsync(EpplusHtmlWriter writer, AccessibilitySettings accessibilitySettings)
        {
            // table header row
            int rowIndex = this._table.Address._toRow;

            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TfootRole))
            {
                writer.AddAttribute("role", accessibilitySettings.TableSettings.TfootRole);
            }

            await writer.RenderBeginTagAsync(HtmlElements.TFoot);
            await writer.ApplyFormatIncreaseIndentAsync(this.Settings.Minify);

            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "row");
                writer.AddAttribute("scope", "row");
            }

            if (this.Settings.SetRowHeight)
            {
                AddRowHeightStyle(writer, this._table.Range, rowIndex, this.Settings.StyleClassPrefix, false);
            }

            await writer.RenderBeginTagAsync(HtmlElements.TableRow);
            await writer.ApplyFormatIncreaseIndentAsync(this.Settings.Minify);
            ExcelAddressBase? address = this._table.Address;
            HtmlImage image = null;

            foreach (int col in this._columns)
            {
                ExcelRange? cell = this._table.WorkSheet.Cells[rowIndex, col];

                if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "cell");
                }

                string? imageCellClassName = GetImageCellClassName(image, this.Settings);
                writer.SetClassAttributeFromStyle(cell, false, this.Settings, imageCellClassName);
                await writer.RenderBeginTagAsync(HtmlElements.TableData);
                await AddImageAsync(writer, this.Settings, image, cell.Value);
                await writer.WriteAsync(GetCellText(cell, this.Settings));
                await writer.RenderEndTagAsync();
                await writer.ApplyFormatAsync(this.Settings.Minify);
            }

            writer.Indent--;
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatDecreaseIndentAsync(this.Settings.Minify);
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatAsync(this.Settings.Minify);
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public async Task<string> GetHtmlStringAsync()
        {
            using MemoryStream? ms = RecyclableMemory.GetStream();
            await this.RenderHtmlAsync(ms);
            ms.Position = 0;
            using StreamReader? sr = new StreamReader(ms);

            return sr.ReadToEnd();
        }

        /// <summary>
        /// Exports the html part of an <see cref="ExcelTable"/> to a stream
        /// </summary>
        /// <returns>A html table</returns>
        public async Task RenderHtmlAsync(Stream stream)
        {
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }

            this.GetDataTypes(this._table.Address, this._table);

            EpplusHtmlWriter? writer = new EpplusHtmlWriter(stream, this.Settings.Encoding, this._styleCache);
            HtmlExportTableUtil.AddClassesAttributes(writer, this._table, this._settings);
            AddTableAccessibilityAttributes(this.Settings.Accessibility, writer);
            await writer.RenderBeginTagAsync(HtmlElements.Table);

            await writer.ApplyFormatIncreaseIndentAsync(this.Settings.Minify);
            this.LoadVisibleColumns();

            if (this.Settings.SetColumnWidth || this.Settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
            {
                await this.SetColumnGroupAsync(writer, this._table.Range, this.Settings, false);
            }

            if (this._table.ShowHeader)
            {
                await this.RenderHeaderRowAsync(writer, this.Settings.Accessibility);
            }

            // table rows
            await this.RenderTableRowsAsync(writer, this.Settings.Accessibility);

            if (this._table.ShowTotal)
            {
                await this.RenderTotalRowAsync(writer, this.Settings.Accessibility);
            }

            // end tag table
            await writer.RenderEndTagAsync();
        }

        /// <summary>
        /// Renders the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public async Task<string> GetSinglePageAsync(string htmlDocument =
                                                         "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            if (this.Settings.Minify)
            {
                htmlDocument = htmlDocument.Replace("\r\n", "");
            }

            string? html = await this.GetHtmlStringAsync();
            CssTableExporterAsync? cssExporter = HtmlExporterFactory.CreateCssExporterTableAsync(this._settings, this._table, this._styleCache);
            string? css = await cssExporter.GetCssStringAsync();

            return string.Format(htmlDocument, html, css);
        }
    }
}
#endif