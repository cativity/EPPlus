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

using OfficeOpenXml.Utils;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Sorting;

namespace OfficeOpenXml.Export.HtmlExport.Exporters;

internal class HtmlTableExporterSync : HtmlRangeExporterSyncBase
{
    internal HtmlTableExporterSync(HtmlTableExportSettings settings, ExcelTable table)
        : base(settings, table.Range)
    {
        Require.Argument(table).IsNotNull("table");
        this._table = table;
        this._tableExportSettings = settings;

        this.LoadRangeImages(new List<ExcelRangeBase>() { table.Range });
    }

    private readonly ExcelTable _table;
    private readonly HtmlTableExportSettings _tableExportSettings;

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

    private void RenderHeaderRow(EpplusHtmlWriter writer)
    {
        // table header row
        if (this.Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(this.Settings.Accessibility.TableSettings.TheadRole))
        {
            writer.AddAttribute("role", this.Settings.Accessibility.TableSettings.TheadRole);
        }

        writer.RenderBeginTag(HtmlElements.Thead);
        writer.ApplyFormatIncreaseIndent(this.Settings.Minify);

        if (this.Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
        {
            writer.AddAttribute("role", "row");
        }

        ExcelAddressBase? adr = this._table.Address;
        int row = adr._fromRow;

        if (this.Settings.SetRowHeight)
        {
            AddRowHeightStyle(writer, this._table.Range, row, this.Settings.StyleClassPrefix, false);
        }

        writer.RenderBeginTag(HtmlElements.TableRow);
        writer.ApplyFormatIncreaseIndent(this.Settings.Minify);
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

            if (this.Settings.Accessibility.TableSettings.AddAccessibilityAttributes
                && !string.IsNullOrEmpty(this.Settings.Accessibility.TableSettings.TableHeaderCellRole))
            {
                writer.AddAttribute("role", this.Settings.Accessibility.TableSettings.TableHeaderCellRole);

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

            writer.RenderBeginTag(HtmlElements.TableHeader);

            if (this.Settings.Pictures.Include == ePictureInclude.Include)
            {
                image = this.GetImage(cell.Worksheet.PositionId, cell._fromRow, cell._fromCol);
            }

            AddImage(writer, this.Settings, image, cell.Value);

            if (cell.Hyperlink == null)
            {
                writer.Write(GetCellText(cell, this.Settings));
            }
            else
            {
                RenderHyperlink(writer, cell, this.Settings);
            }

            writer.RenderEndTag();
            writer.ApplyFormat(this.Settings.Minify);
        }

        writer.Indent--;
        writer.RenderEndTag();
        writer.ApplyFormatDecreaseIndent(this.Settings.Minify);
        writer.RenderEndTag();
        writer.ApplyFormat(this.Settings.Minify);
    }

    private void RenderTableRows(EpplusHtmlWriter writer, AccessibilitySettings accessibilitySettings)
    {
        if (accessibilitySettings.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(accessibilitySettings.TableSettings.TbodyRole))
        {
            writer.AddAttribute("role", accessibilitySettings.TableSettings.TbodyRole);
        }

        writer.RenderBeginTag(HtmlElements.Tbody);
        writer.ApplyFormatIncreaseIndent(this.Settings.Minify);
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

            writer.RenderBeginTag(HtmlElements.TableRow);
            writer.ApplyFormatIncreaseIndent(this.Settings.Minify);

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

                    CellDataWriter.Write(cell, dataType, writer, this.Settings, accessibilitySettings, addRowScope, image);
                }
                else
                {
                    writer.RenderBeginTag(HtmlElements.TableData);
                    AddImage(writer, this.Settings, image, cell.Value);
                    string? imageCellClassName = GetImageCellClassName(image, this.Settings);
                    writer.SetClassAttributeFromStyle(cell, false, this.Settings, imageCellClassName);
                    RenderHyperlink(writer, cell, this.Settings);
                    writer.RenderEndTag();
                    writer.ApplyFormat(this.Settings.Minify);
                }
            }

            // end tag tr
            writer.Indent--;
            writer.RenderEndTag();
            writer.ApplyFormat(this.Settings.Minify);
            row++;
        }

        writer.ApplyFormatDecreaseIndent(this.Settings.Minify);

        // end tag tbody
        writer.RenderEndTag();
        writer.ApplyFormat(this.Settings.Minify);
    }

    private void RenderTotalRow(EpplusHtmlWriter writer)
    {
        // table header row
        int rowIndex = this._table.Address._toRow;

        if (this.Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(this.Settings.Accessibility.TableSettings.TfootRole))
        {
            writer.AddAttribute("role", this.Settings.Accessibility.TableSettings.TfootRole);
        }

        writer.RenderBeginTag(HtmlElements.TFoot);
        writer.ApplyFormatIncreaseIndent(this.Settings.Minify);

        if (this.Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
        {
            writer.AddAttribute("role", "row");
            writer.AddAttribute("scope", "row");
        }

        if (this.Settings.SetRowHeight)
        {
            AddRowHeightStyle(writer, this._table.Range, rowIndex, this.Settings.StyleClassPrefix, false);
        }

        writer.RenderBeginTag(HtmlElements.TableRow);
        writer.ApplyFormatIncreaseIndent(this.Settings.Minify);
        HtmlImage image = null;

        foreach (int col in this._columns)
        {
            ExcelRange? cell = this._table.WorkSheet.Cells[rowIndex, col];

            if (this.Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "cell");
            }

            string? imageCellClassName = GetImageCellClassName(image, this.Settings);
            writer.SetClassAttributeFromStyle(cell, false, this.Settings, imageCellClassName);
            writer.RenderBeginTag(HtmlElements.TableData);
            AddImage(writer, this.Settings, image, cell.Value);
            writer.Write(GetCellText(cell, this.Settings));
            writer.RenderEndTag();
            writer.ApplyFormat(this.Settings.Minify);
        }

        writer.Indent--;
        writer.RenderEndTag();
        writer.ApplyFormatDecreaseIndent(this.Settings.Minify);
        writer.RenderEndTag();
        writer.ApplyFormat(this.Settings.Minify);
    }

    /// <summary>
    /// Exports an <see cref="ExcelTable"/> to a html string
    /// </summary>
    /// <returns>A html table</returns>
    public string GetHtmlString()
    {
        using MemoryStream? ms = RecyclableMemory.GetStream();
        this.RenderHtml(ms);
        ms.Position = 0;
        using StreamReader? sr = new StreamReader(ms);

        return sr.ReadToEnd();
    }

    /// <summary>
    /// Exports the html part of an <see cref="ExcelTable"/> to a html string.
    /// </summary>
    /// <param name="stream">The stream to write to.</param>
    /// <exception cref="IOException"></exception>
    public void RenderHtml(Stream stream)
    {
        if (!stream.CanWrite)
        {
            throw new IOException("Parameter stream must be a writeable System.IO.Stream");
        }

        this.GetDataTypes(this._table.Address, this._table);

        EpplusHtmlWriter? writer = new EpplusHtmlWriter(stream, this.Settings.Encoding, this._styleCache);
        HtmlExportTableUtil.AddClassesAttributes(writer, this._table, this._tableExportSettings);
        AddTableAccessibilityAttributes(this.Settings.Accessibility, writer);
        writer.RenderBeginTag(HtmlElements.Table);

        writer.ApplyFormatIncreaseIndent(this.Settings.Minify);
        this.LoadVisibleColumns();

        if (this.Settings.SetColumnWidth || this.Settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.ColumnDataType)
        {
            this.SetColumnGroup(writer, this._table.Range, this.Settings, false);
        }

        if (this._table.ShowHeader)
        {
            this.RenderHeaderRow(writer);
        }

        // table rows
        this.RenderTableRows(writer, this.Settings.Accessibility);

        if (this._table.ShowTotal)
        {
            this.RenderTotalRow(writer);
        }

        // end tag table
        writer.RenderEndTag();
    }

    /// <summary>
    /// Renders both the Css and the Html to a single page. 
    /// </summary>
    /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
    /// <returns>The html document</returns>
    public string GetSinglePage(string htmlDocument =
                                    "<!DOCTYPE html>\r\n<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
    {
        if (this.Settings.Minify)
        {
            htmlDocument = htmlDocument.Replace("\r\n", "");
        }

        string? html = this.GetHtmlString();
        CssTableExporterSync? cssExporter = HtmlExporterFactory.CreateCssExporterTableSync(this._tableExportSettings, this._table, this._styleCache);
        string? css = cssExporter.GetCssString();

        return string.Format(htmlDocument, html, css);
    }
}