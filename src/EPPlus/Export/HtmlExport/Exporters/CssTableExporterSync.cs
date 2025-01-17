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

namespace OfficeOpenXml.Export.HtmlExport.Exporters;

internal class CssTableExporterSync : CssRangeExporterBase
{
    public CssTableExporterSync(HtmlTableExportSettings settings, ExcelTable table)
        : base(settings, table.Range)
    {
        this._table = table;
        this._tableSettings = settings;
    }

    private readonly ExcelTable _table;
    private readonly HtmlTableExportSettings _tableSettings;

    private static void RenderTableCss(StreamWriter sw,
                                       ExcelTable table,
                                       HtmlTableExportSettings settings,
                                       Dictionary<string, int> styleCache,
                                       List<string> datatypes)
    {
        EpplusTableCssWriter? styleWriter = new EpplusTableCssWriter(sw, table, settings, styleCache);

        if (settings.Minify == false)
        {
            styleWriter.WriteLine();
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
            $"{TableClass}.{HtmlExportTableUtil.TableStyleClassPrefix}{HtmlExportTableUtil.GetClassName(tblStyle.Name, "EmptyTableStyle").ToLower()}";

        styleWriter.AddHyperlinkCss($"{tableClass}", tblStyle.WholeTable);
        styleWriter.AddAlignmentToCss($"{tableClass}", datatypes);

        styleWriter.AddToCss($"{tableClass}", tblStyle.WholeTable, "");
        styleWriter.AddToCssBorderVH($"{tableClass}", tblStyle.WholeTable, "");

        //Header
        styleWriter.AddToCss($"{tableClass}", tblStyle.HeaderRow, " thead");
        styleWriter.AddToCssBorderVH($"{tableClass}", tblStyle.HeaderRow, "");

        styleWriter.AddToCss($"{tableClass}", tblStyle.LastTotalCell, $" thead tr th:last-child)");
        styleWriter.AddToCss($"{tableClass}", tblStyle.FirstHeaderCell, " thead tr th:first-child");

        //Total
        styleWriter.AddToCss($"{tableClass}", tblStyle.TotalRow, " tfoot");
        styleWriter.AddToCssBorderVH($"{tableClass}", tblStyle.TotalRow, "");
        styleWriter.AddToCss($"{tableClass}", tblStyle.LastTotalCell, $" tfoot tr td:last-child)");
        styleWriter.AddToCss($"{tableClass}", tblStyle.FirstTotalCell, " tfoot tr td:first-child");

        //Columns stripes
        string? tableClassCS = $"{tableClass}-column-stripes";
        styleWriter.AddToCss($"{tableClassCS}", tblStyle.FirstColumnStripe, $" tbody tr td:nth-child(odd)");
        styleWriter.AddToCss($"{tableClassCS}", tblStyle.SecondColumnStripe, $" tbody tr td:nth-child(even)");

        //Row stripes
        string? tableClassRS = $"{tableClass}-row-stripes";
        styleWriter.AddToCss($"{tableClassRS}", tblStyle.FirstRowStripe, " tbody tr:nth-child(odd)");
        styleWriter.AddToCss($"{tableClassRS}", tblStyle.SecondRowStripe, " tbody tr:nth-child(even)");

        //Last column
        string? tableClassLC = $"{tableClass}-last-column";
        styleWriter.AddToCss($"{tableClassLC}", tblStyle.LastColumn, $" tbody tr td:last-child");

        //First column
        string? tableClassFC = $"{tableClass}-first-column";
        styleWriter.AddToCss($"{tableClassFC}", tblStyle.FirstColumn, " tbody tr td:first-child");

        styleWriter.FlushStream();
    }

    private void RenderCellCss(EpplusCssWriter styleWriter)
    {
        ExcelRangeBase? r = this._table.Range;
        ExcelStyles? styles = r.Worksheet.Workbook.Styles;
        CellStoreEnumerator<ExcelValue>? ce = new CellStoreEnumerator<ExcelValue>(r.Worksheet._values, r._fromRow, r._fromCol, r._toRow, r._toCol);

        while (ce.Next())
        {
            if (ce.Value._styleId > 0 && ce.Value._styleId < styles.CellXfs.Count)
            {
                styleWriter.AddToCss(styles, ce.Value._styleId, this.Settings.StyleClassPrefix, this.Settings.CellStyleClassName);
            }
        }

        styleWriter.FlushStream();
    }

    /// <summary>
    /// Exports an <see cref="ExcelTable"/> to a html string
    /// </summary>
    /// <returns>A html table</returns>
    public string GetCssString()
    {
        using MemoryStream? ms = RecyclableMemory.GetStream();
        this.RenderCss(ms);
        ms.Position = 0;
        using StreamReader? sr = new StreamReader(ms);

        return sr.ReadToEnd();
    }

    /// <summary>
    /// Exports the css part of an <see cref="ExcelTable"/> to a html string
    /// </summary>
    /// <returns>A html table</returns>
    public void RenderCss(Stream stream)
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

        cellCssWriter.RenderAdditionalAndFontCss(TableClass);

        if (this._tableSettings.Css.IncludeTableStyles)
        {
            RenderTableCss(sw, this._table, this._tableSettings, this._styleCache, this._dataTypes);
        }

        if (this._tableSettings.Css.IncludeCellStyles)
        {
            this.RenderCellCss(cellCssWriter);
        }

        if (this._tableSettings.Pictures.Include == ePictureInclude.Include)
        {
            this.LoadRangeImages(ranges);

            foreach (HtmlImage? p in this._rangePictures)
            {
                cellCssWriter.AddPictureToCss(p);
            }
        }

        cellCssWriter.FlushStream();
    }
}