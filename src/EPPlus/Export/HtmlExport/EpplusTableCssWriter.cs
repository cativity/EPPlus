/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2021         EPPlus Software AB       Added Html Export
 *************************************************************************************************/

using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Table;
using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Table;
using System.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style.Dxf;
using static OfficeOpenXml.Export.HtmlExport.ColumnDataTypeManager;
using System.Text;
using System.Globalization;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Export.HtmlExport.Settings;

namespace OfficeOpenXml.Export.HtmlExport;

internal partial class EpplusTableCssWriter : HtmlWriterBase
{
    protected HtmlTableExportSettings _settings;
    private readonly List<EpplusHtmlAttribute> _attributes = new List<EpplusHtmlAttribute>();
    ExcelTable _table;
    ExcelTheme _theme;

    internal EpplusTableCssWriter(Stream stream, ExcelTable table, HtmlTableExportSettings settings, Dictionary<string, int> styleCache)
        : base(stream, settings.Encoding, styleCache)
    {
        this.Init(table, settings);
    }

    internal EpplusTableCssWriter(StreamWriter writer, ExcelTable table, HtmlTableExportSettings settings, Dictionary<string, int> styleCache)
        : base(writer, styleCache)
    {
        this.Init(table, settings);
    }

    private void Init(ExcelTable table, HtmlTableExportSettings settings)
    {
        this._table = table;
        this._settings = settings;

        if (table.WorkSheet.Workbook.ThemeManager.CurrentTheme == null)
        {
            table.WorkSheet.Workbook.ThemeManager.CreateDefaultTheme();
        }

        this._theme = table.WorkSheet.Workbook.ThemeManager.CurrentTheme;
    }

    internal void AddAlignmentToCss(string name, List<string> dataTypes)
    {
        if (this._settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.DontSet)
        {
            return;
        }

        int row = this._table.ShowHeader ? this._table.Address._fromRow + 1 : this._table.Address._fromRow;

        for (int c = 0; c < this._table.Columns.Count; c++)
        {
            int col = this._table.Address._fromCol + c;
            int styleId = this._table.WorkSheet.GetStyleInner(row, col);
            string hAlign = "";
            string vAlign = "";

            if (styleId > 0)
            {
                ExcelXfs? xfs = this._table.WorkSheet.Workbook.Styles.CellXfs[styleId];

                if (xfs.ApplyAlignment ?? true)
                {
                    hAlign = GetHorizontalAlignment(xfs);
                    vAlign = GetVerticalAlignment(xfs);
                }
            }

            if (string.IsNullOrEmpty(hAlign) && c < dataTypes.Count && (dataTypes[c] == HtmlDataTypes.Number || dataTypes[c] == HtmlDataTypes.DateTime))
            {
                hAlign = "right";
            }

            if (!(string.IsNullOrEmpty(hAlign) && string.IsNullOrEmpty(vAlign)))
            {
                this.WriteClass($"table.{name} td:nth-child({col}){{", this._settings.Minify);

                if (string.IsNullOrEmpty(hAlign) == false && this._settings.Css.Exclude.TableStyle.HorizontalAlignment == false)
                {
                    this.WriteCssItem($"text-align:{hAlign};", this._settings.Minify);
                }

                if (string.IsNullOrEmpty(vAlign) == false && this._settings.Css.Exclude.TableStyle.VerticalAlignment == false)
                {
                    this.WriteCssItem($"vertical-align:{vAlign};", this._settings.Minify);
                }

                this.WriteClassEnd(this._settings.Minify);
            }
        }
    }

    internal void AddToCss(string name, ExcelTableStyleElement element, string htmlElement)
    {
        ExcelDxfStyleLimitedFont? s = element.Style;

        if (s.HasValue == false)
        {
            return; //Dont add empty elements
        }

        this.WriteClass($"table.{name}{htmlElement}{{", this._settings.Minify);
        this.WriteFillStyles(s.Fill);
        this.WriteFontStyles(s.Font);
        this.WriteBorderStyles(s.Border);
        this.WriteClassEnd(this._settings.Minify);
    }

    internal void AddHyperlinkCss(string name, ExcelTableStyleElement element)
    {
        this.WriteClass($"table.{name} a{{", this._settings.Minify);
        this.WriteFontStyles(element.Style.Font);
        this.WriteClassEnd(this._settings.Minify);
    }

    internal void AddToCssBorderVH(string name, ExcelTableStyleElement element, string htmlElement)
    {
        ExcelDxfStyleLimitedFont? s = element.Style;

        if (s.Border.Vertical.HasValue == false && s.Border.Horizontal.HasValue == false)
        {
            return; //Dont add empty elements
        }

        this.WriteClass($"table.{name}{htmlElement} td,tr {{", this._settings.Minify);
        this.WriteBorderStylesVerticalHorizontal(s.Border);
        this.WriteClassEnd(this._settings.Minify);
    }

    internal void FlushStream()
    {
        this._writer.Flush();
    }

    private void WriteFillStyles(ExcelDxfFill f)
    {
        if (f.HasValue && this._settings.Css.Exclude.TableStyle.Fill == false)
        {
            if (f.Style == eDxfFillStyle.PatternFill)
            {
                if (f.PatternType.Value == ExcelFillStyle.Solid)
                {
                    this.WriteCssItem($"background-color:{this.GetDxfColor(f.BackgroundColor)};", this._settings.Minify);
                }
                else
                {
                    this.WriteCssItem($"{PatternFills.GetPatternSvg(f.PatternType.Value, this.GetDxfColor(f.BackgroundColor), this.GetDxfColor(f.PatternColor))};",
                                      this._settings.Minify);
                }
            }
            else if (f.Style == eDxfFillStyle.GradientFill)
            {
                this.WriteDxfGradient(f.Gradient);
            }
        }
    }

    private void WriteDxfGradient(ExcelDxfGradientFill gradient)
    {
        StringBuilder? sb = new StringBuilder();

        if (gradient.GradientType == eDxfGradientFillType.Linear)
        {
            _ = sb.Append($"background: linear-gradient({(gradient.Degree + 90) % 360}deg");
        }
        else
        {
            _ = sb.Append($"background:radial-gradient(ellipse {(gradient.Right ?? 0) * 100}% {(gradient.Bottom ?? 0) * 100}%");
        }

        foreach (ExcelDxfGradientFillColor? color in gradient.Colors)
        {
            _ = sb.Append($",{this.GetDxfColor(color.Color)} {color.Position.ToString("F", CultureInfo.InvariantCulture)}%");
        }

        _ = sb.Append(")");

        this.WriteCssItem(sb.ToString(), this._settings.Minify);
    }

    private void WriteFontStyles(ExcelDxfFontBase f)
    {
        eFontExclude flags = this._settings.Css.Exclude.TableStyle.Font;

        if (f.Color.HasValue && EnumUtil.HasNotFlag(flags, eFontExclude.Color))
        {
            this.WriteCssItem($"color:{this.GetDxfColor(f.Color)};", this._settings.Minify);
        }

        if (f.Bold.HasValue && f.Bold.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Bold))
        {
            this.WriteCssItem("font-weight:bolder;", this._settings.Minify);
        }

        if (f.Italic.HasValue && f.Italic.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Italic))
        {
            this.WriteCssItem("font-style:italic;", this._settings.Minify);
        }

        if (f.Strike.HasValue && f.Strike.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Strike))
        {
            this.WriteCssItem("text-decoration:line-through solid;", this._settings.Minify);
        }

        if (f.Underline.HasValue && f.Underline != ExcelUnderLineType.None && EnumUtil.HasNotFlag(flags, eFontExclude.Underline))
        {
            this.WriteCssItem("text-decoration:underline ", this._settings.Minify);

            switch (f.Underline.Value)
            {
                case ExcelUnderLineType.Double:
                case ExcelUnderLineType.DoubleAccounting:
                    this.WriteCssItem("double;", this._settings.Minify);

                    break;

                default:
                    this.WriteCssItem("solid;", this._settings.Minify);

                    break;
            }
        }
    }

    private void WriteBorderStyles(ExcelDxfBorderBase b)
    {
        if (b.HasValue)
        {
            eBorderExclude flags = this._settings.Css.Exclude.TableStyle.Border;

            if (EnumUtil.HasNotFlag(flags, eBorderExclude.Top))
            {
                this.WriteBorderItem(b.Top, "top");
            }

            if (EnumUtil.HasNotFlag(flags, eBorderExclude.Bottom))
            {
                this.WriteBorderItem(b.Bottom, "bottom");
            }

            if (EnumUtil.HasNotFlag(flags, eBorderExclude.Left))
            {
                this.WriteBorderItem(b.Left, "left");
            }

            if (EnumUtil.HasNotFlag(flags, eBorderExclude.Right))
            {
                this.WriteBorderItem(b.Right, "right");
            }
        }
    }

    private void WriteBorderStylesVerticalHorizontal(ExcelDxfBorderBase b)
    {
        if (b.HasValue)
        {
            eBorderExclude flags = this._settings.Css.Exclude.TableStyle.Border;

            if (EnumUtil.HasNotFlag(flags, eBorderExclude.Top))
            {
                this.WriteBorderItem(b.Horizontal, "top");
            }

            if (EnumUtil.HasNotFlag(flags, eBorderExclude.Bottom))
            {
                this.WriteBorderItem(b.Horizontal, "bottom");
            }

            if (EnumUtil.HasNotFlag(flags, eBorderExclude.Left))
            {
                this.WriteBorderItem(b.Vertical, "left");
            }

            if (EnumUtil.HasNotFlag(flags, eBorderExclude.Right))
            {
                this.WriteBorderItem(b.Vertical, "right");
            }
        }
    }

    private void WriteBorderItem(ExcelDxfBorderItem bi, string suffix)
    {
        if (bi.HasValue && bi.Style != ExcelBorderStyle.None)
        {
            StringBuilder? sb = new StringBuilder();
            _ = sb.Append(GetBorderItemLine(bi.Style.Value, suffix));

            if (bi.Color.HasValue)
            {
                _ = sb.Append($" {this.GetDxfColor(bi.Color)}");
            }

            _ = sb.Append(";");

            this.WriteCssItem(sb.ToString(), this._settings.Minify);
        }
    }

    private string GetDxfColor(ExcelDxfColor c)
    {
        Color ret;

        if (c.Color.HasValue)
        {
            ret = c.Color.Value;
        }
        else if (c.Theme.HasValue)
        {
            ret = Utils.ColorConverter.GetThemeColor(this._theme, c.Theme.Value);
        }
        else if (c.Index.HasValue)
        {
            ret = ExcelColor.GetIndexedColor(c.Index.Value);
        }
        else
        {
            //Automatic, set to black.
            ret = Color.Black;
        }

        if (c.Tint.HasValue)
        {
            ret = Utils.ColorConverter.ApplyTint(ret, c.Tint.Value);
        }

        return "#" + ret.ToArgb().ToString("x8").Substring(2);
    }
}