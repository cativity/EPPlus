/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2022         EPPlus Software AB       EPPlus 6
 *************************************************************************************************/

using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using static OfficeOpenXml.ExcelAddressBase;

namespace OfficeOpenXml.Core;

internal class AutofitHelper
{
    private ExcelRangeBase _range;
    ITextMeasurer _genericMeasurer = new GenericFontMetricsTextMeasurer();
    MeasurementFont _nonExistingFont = new MeasurementFont() { FontFamily = FontSize.NonExistingFont };
    Dictionary<float, short> _fontWidthDefault = null;
    Dictionary<int, MeasurementFont> _fontCache;

    public AutofitHelper(ExcelRangeBase range)
    {
        this._range = range;

        if (FontSize.FontWidths.ContainsKey(FontSize.NonExistingFont))
        {
            FontSize.LoadAllFontsFromResource();
            this._fontWidthDefault = FontSize.FontWidths[FontSize.NonExistingFont];
        }
    }

    internal void AutofitColumn(double MinimumWidth, double MaximumWidth)
    {
        ExcelWorksheet? ws = this._range._worksheet;

        if (ws.Dimension == null)
        {
            return;
        }

        if (this._range._fromCol < 1 || this._range._fromRow < 1)
        {
            this._range.SetToSelectedRange();
        }

        this._fontCache = new Dictionary<int, MeasurementFont>();

        bool doAdjust = ws._package.DoAdjustDrawings;
        ws._package.DoAdjustDrawings = false;
        double[,]? drawWidths = ws.Drawings.GetDrawingWidths();

        int fromCol = this._range._fromCol > ws.Dimension._fromCol ? this._range._fromCol : ws.Dimension._fromCol;
        int toCol = this._range._toCol < ws.Dimension._toCol ? this._range._toCol : ws.Dimension._toCol;

        if (fromCol > toCol)
        {
            return; //Issue 15383
        }

        if (this._range.Addresses == null)
        {
            SetMinWidth(ws, MinimumWidth, fromCol, toCol);
        }
        else
        {
            foreach (ExcelAddressBase? addr in this._range.Addresses)
            {
                fromCol = addr._fromCol > ws.Dimension._fromCol ? addr._fromCol : ws.Dimension._fromCol;
                toCol = addr._toCol < ws.Dimension._toCol ? addr._toCol : ws.Dimension._toCol;
                SetMinWidth(ws, MinimumWidth, fromCol, toCol);
            }
        }

        //Get any autofilter to widen these columns
        List<ExcelAddressBase>? afAddr = new List<ExcelAddressBase>();

        if (ws.AutoFilterAddress != null)
        {
            afAddr.Add(new ExcelAddressBase(ws.AutoFilterAddress._fromRow,
                                            ws.AutoFilterAddress._fromCol,
                                            ws.AutoFilterAddress._fromRow,
                                            ws.AutoFilterAddress._toCol));

            afAddr[afAddr.Count - 1]._ws = this._range.WorkSheetName;
        }

        foreach (ExcelTable? tbl in ws.Tables)
        {
            if (tbl.AutoFilterAddress != null)
            {
                afAddr.Add(new ExcelAddressBase(tbl.AutoFilterAddress._fromRow,
                                                tbl.AutoFilterAddress._fromCol,
                                                tbl.AutoFilterAddress._fromRow,
                                                tbl.AutoFilterAddress._toCol));

                afAddr[afAddr.Count - 1]._ws = this._range.WorkSheetName;
            }
        }

        ExcelStyles? styles = ws.Workbook.Styles;
        ExcelNamedStyleXml? ns = styles.GetNormalStyle();
        int normalXfId = ns?.StyleXfId ?? 0;

        if (normalXfId < 0 || normalXfId >= styles.CellStyleXfs.Count)
        {
            normalXfId = 0;
        }

        ExcelFontXml? nf = styles.Fonts[styles.CellStyleXfs[normalXfId].FontId];
        MeasurementFontStyles fs;
        //MeasurementFontStyles fs = MeasurementFontStyles.Regular;

        //if (nf.Bold)
        //{
        //    fs |= MeasurementFontStyles.Bold;
        //}

        //if (nf.UnderLine)
        //{
        //    fs |= MeasurementFontStyles.Underline;
        //}

        //if (nf.Italic)
        //{
        //    fs |= MeasurementFontStyles.Italic;
        //}

        //if (nf.Strike)
        //{
        //    fs |= MeasurementFontStyles.Strikeout;
        //}

        float normalSize = Convert.ToSingle(FontSize.GetWidthPixels(nf.Name, nf.Size));
        ExcelTextSettings? textSettings = this._range._workbook._package.Settings.TextSettings;

        foreach (ExcelRangeBase? cell in this._range)
        {
            if (ws.Column(cell.Start.Column).Hidden) //Issue 15338
            {
                continue;
            }

            if (cell.Merge == true || styles.CellXfs[cell.StyleID].WrapText)
            {
                continue;
            }

            int fntID = styles.CellXfs[cell.StyleID].FontId;
            MeasurementFont f;

            if (this._fontCache.ContainsKey(fntID))
            {
                f = this._fontCache[fntID];
            }
            else
            {
                ExcelFontXml? fnt = styles.Fonts[fntID];
                fs = MeasurementFontStyles.Regular;

                if (fnt.Bold)
                {
                    fs |= MeasurementFontStyles.Bold;
                }

                if (fnt.UnderLine)
                {
                    fs |= MeasurementFontStyles.Underline;
                }

                if (fnt.Italic)
                {
                    fs |= MeasurementFontStyles.Italic;
                }

                if (fnt.Strike)
                {
                    fs |= MeasurementFontStyles.Strikeout;
                }

                f = new MeasurementFont { FontFamily = fnt.Name, Style = fs, Size = fnt.Size };

                this._fontCache.Add(fntID, f);
            }

            int ind = styles.CellXfs[cell.StyleID].Indent;
            string? textForWidth = cell.TextForWidth;
            string? t = textForWidth + (ind > 0 && !string.IsNullOrEmpty(textForWidth) ? new string('_', ind) : "");

            if (t.Length > 32000)
            {
                t = t.Substring(0, 32000); //Issue
            }

            TextMeasurement size = this.MeasureString(t, fntID, textSettings);

            double width;
            double r = styles.CellXfs[cell.StyleID].TextRotation;

            if (r <= 0)
            {
                int padding = 0; // 5
                width = (size.Width + padding) / normalSize;
            }
            else
            {
                r = r <= 90 ? r : r - 90;
                width = (((size.Width - size.Height) * Math.Abs(Math.Cos(Math.PI * r / 180.0))) + size.Height + 5) / normalSize;
            }

            foreach (ExcelAddressBase? a in afAddr)
            {
                if (a.Collide(cell) != eAddressCollition.No)
                {
                    width += 2.25;

                    break;
                }
            }

            if (width > ws.Column(cell._fromCol).Width)
            {
                ws.Column(cell._fromCol).Width = width > MaximumWidth ? MaximumWidth : width;
            }
        }

        ws.Drawings.AdjustWidth(drawWidths);
        ws._package.DoAdjustDrawings = doAdjust;
    }

    private TextMeasurement MeasureString(string t, int fntID, ExcelTextSettings ts)
    {
        Dictionary<ulong, TextMeasurement>? measureCache = new Dictionary<ulong, TextMeasurement>();
        ulong key = ((ulong)(uint)t.GetHashCode() << 32) | (uint)fntID;

        if (!measureCache.TryGetValue(key, out TextMeasurement measurement))
        {
            ITextMeasurer? measurer = ts.PrimaryTextMeasurer;
            MeasurementFont? font = this._fontCache[fntID];
            measurement = measurer.MeasureText(t, font);

            if (measurement.IsEmpty && ts.FallbackTextMeasurer != null && ts.FallbackTextMeasurer != ts.PrimaryTextMeasurer)
            {
                measurer = ts.FallbackTextMeasurer;
                measurement = measurer.MeasureText(t, font);
            }

            if (measurement.IsEmpty && this._fontWidthDefault != null)
            {
                measurement = this.MeasureGeneric(t, ts, font);
            }

            if (!measurement.IsEmpty && ts.AutofitScaleFactor != 1f)
            {
                measurement.Height *= ts.AutofitScaleFactor;
                measurement.Width *= ts.AutofitScaleFactor;
            }

            measureCache.Add(key, measurement);
        }

        return measurement;
    }

    private TextMeasurement MeasureGeneric(string t, ExcelTextSettings ts, MeasurementFont font)
    {
        TextMeasurement measurement;

        if (FontSize.FontWidths.ContainsKey(font.FontFamily))
        {
            decimal width = FontSize.GetWidthPixels(font.FontFamily, font.Size);
            decimal height = FontSize.GetHeightPixels(font.FontFamily, font.Size);
            decimal defaultWidth = FontSize.GetWidthPixels(FontSize.NonExistingFont, font.Size);
            decimal defaultHeight = FontSize.GetHeightPixels(FontSize.NonExistingFont, font.Size);
            this._nonExistingFont.Size = font.Size;
            this._nonExistingFont.Style = font.Style;
            measurement = this._genericMeasurer.MeasureText(t, this._nonExistingFont);

            measurement.Width *= (float)(width / defaultWidth) * ts.AutofitScaleFactor;
            measurement.Height *= (float)(height / defaultHeight) * ts.AutofitScaleFactor;
        }
        else
        {
            this._nonExistingFont.Size = font.Size;
            this._nonExistingFont.Style = font.Style;
            measurement = this._genericMeasurer.MeasureText(t, this._nonExistingFont);
            measurement.Height *= ts.AutofitScaleFactor;
            measurement.Width *= ts.AutofitScaleFactor;
        }

        return measurement;
    }

    private static void SetMinWidth(ExcelWorksheet ws, double minimumWidth, int fromCol, int toCol)
    {
        CellStoreEnumerator<ExcelValue>? iterator = new CellStoreEnumerator<ExcelValue>(ws._values, 0, fromCol, 0, toCol);
        int prevCol = fromCol;

        foreach (ExcelValue val in iterator)
        {
            ExcelColumn? col = (ExcelColumn)val._value;

            if (col.Hidden)
            {
                continue;
            }

            col.Width = minimumWidth;

            if (ws.DefaultColWidth > minimumWidth && col.ColumnMin > prevCol)
            {
                ExcelColumn? newCol = ws.Column(prevCol);
                newCol.ColumnMax = col.ColumnMin - 1;
                newCol.Width = minimumWidth;
            }

            prevCol = col.ColumnMax + 1;
        }

        if (ws.DefaultColWidth > minimumWidth && prevCol < toCol)
        {
            ExcelColumn? newCol = ws.Column(prevCol);
            newCol.ColumnMax = toCol;
            newCol.Width = minimumWidth;
        }
    }
}