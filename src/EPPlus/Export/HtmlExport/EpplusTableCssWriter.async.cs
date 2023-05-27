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
using OfficeOpenXml.Export.HtmlExport.Exporters;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
#if !NET35 && !NET40
    internal partial class EpplusTableCssWriter : HtmlWriterBase
    {
        internal async Task RenderAdditionalAndFontCssAsync()
        {
            await this.WriteClassAsync($"table.{AbstractHtmlExporter.TableClass}{{", this._settings.Minify);
            ExcelNamedStyleXml? ns = this._table.WorkSheet.Workbook.Styles.GetNormalStyle();

            if (ns != null)
            {
                await this.WriteCssItemAsync($"font-family:{ns.Style.Font.Name};", this._settings.Minify);
                await this.WriteCssItemAsync($"font-size:{ns.Style.Font.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", this._settings.Minify);
            }

            foreach (KeyValuePair<string, string> item in this._settings.Css.AdditionalCssElements)
            {
                await this.WriteCssItemAsync($"{item.Key}:{item.Value};", this._settings.Minify);
            }

            await this.WriteClassEndAsync(this._settings.Minify);
        }

        internal async Task AddAlignmentToCssAsync(string name, List<string> dataTypes)
        {
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

                if (string.IsNullOrEmpty(hAlign) && c < dataTypes.Count && dataTypes[c] == HtmlDataTypes.Number)
                {
                    hAlign = "right";
                }

                if (!(string.IsNullOrEmpty(hAlign) && string.IsNullOrEmpty(vAlign)))
                {
                    await this.WriteClassAsync($"table.{name} td:nth-child({col}){{", this._settings.Minify);

                    if (string.IsNullOrEmpty(hAlign) == false && this._settings.Css.Exclude.TableStyle.HorizontalAlignment == false)
                    {
                        await this.WriteCssItemAsync($"text-align:{hAlign};", this._settings.Minify);
                    }

                    if (string.IsNullOrEmpty(vAlign) == false && this._settings.Css.Exclude.TableStyle.VerticalAlignment == false)
                    {
                        await this.WriteCssItemAsync($"vertical-align:{vAlign};", this._settings.Minify);
                    }

                    await this.WriteClassEndAsync(this._settings.Minify);
                }
            }
        }

        internal async Task AddToCssAsync(string name, ExcelTableStyleElement element, string htmlElement)
        {
            ExcelDxfStyleLimitedFont? s = element.Style;

            if (s.HasValue == false)
            {
                return; //Dont add empty elements
            }

            await this.WriteClassAsync($"table.{name}{htmlElement}{{", this._settings.Minify);
            await this.WriteFillStylesAsync(s.Fill);
            await this.WriteFontStylesAsync(s.Font);
            await this.WriteBorderStylesAsync(s.Border);
            await this.WriteClassEndAsync(this._settings.Minify);
        }

        internal async Task AddHyperlinkCssAsync(string name, ExcelTableStyleElement element)
        {
            await this.WriteClassAsync($"table.{name} a{{", this._settings.Minify);
            await this.WriteFontStylesAsync(element.Style.Font);
            await this.WriteClassEndAsync(this._settings.Minify);
        }

        internal async Task AddToCssBorderVHAsync(string name, ExcelTableStyleElement element, string htmlElement)
        {
            ExcelDxfStyleLimitedFont? s = element.Style;

            if (s.Border.Vertical.HasValue == false && s.Border.Horizontal.HasValue == false)
            {
                return; //Dont add empty elements
            }

            await this.WriteClassAsync($"table.{name}{htmlElement} td,tr {{", this._settings.Minify);
            await this.WriteBorderStylesVerticalHorizontalAsync(s.Border);
            await this.WriteClassEndAsync(this._settings.Minify);
        }

        internal async Task FlushStreamAsync()
        {
            await this._writer.FlushAsync();
        }

        private async Task WriteFillStylesAsync(ExcelDxfFill f)
        {
            if (f.HasValue && this._settings.Css.Exclude.TableStyle.Fill == false)
            {
                if (f.Style == eDxfFillStyle.PatternFill)
                {
                    if (f.PatternType.Value == ExcelFillStyle.Solid)
                    {
                        await this.WriteCssItemAsync($"background-color:{this.GetDxfColor(f.BackgroundColor)};", this._settings.Minify);
                    }
                    else
                    {
                        await
                            this.WriteCssItemAsync($"{PatternFills.GetPatternSvg(f.PatternType.Value, this.GetDxfColor(f.BackgroundColor), this.GetDxfColor(f.PatternColor))};",
                                                   this._settings.Minify);
                    }
                }
                else if (f.Style == eDxfFillStyle.GradientFill)
                {
                    await this.WriteDxfGradientAsync(f.Gradient);
                }
            }
        }

        private async Task WriteDxfGradientAsync(ExcelDxfGradientFill gradient)
        {
            StringBuilder? sb = new StringBuilder();

            if (gradient.GradientType == eDxfGradientFillType.Linear)
            {
                sb.Append($"background: linear-gradient({(gradient.Degree + 90) % 360}deg");
            }
            else
            {
                sb.Append($"background:radial-gradient(ellipse {(gradient.Right ?? 0) * 100}% {(gradient.Bottom ?? 0) * 100}%");
            }

            foreach (ExcelDxfGradientFillColor? color in gradient.Colors)
            {
                sb.Append($",{this.GetDxfColor(color.Color)} {color.Position.ToString("F", CultureInfo.InvariantCulture)}%");
            }

            sb.Append(")");

            await this.WriteCssItemAsync(sb.ToString(), this._settings.Minify);
        }

        private async Task WriteFontStylesAsync(ExcelDxfFontBase f)
        {
            eFontExclude flags = this._settings.Css.Exclude.TableStyle.Font;

            if (f.Color.HasValue && EnumUtil.HasNotFlag(flags, eFontExclude.Color))
            {
                await this.WriteCssItemAsync($"color:{this.GetDxfColor(f.Color)};", this._settings.Minify);
            }

            if (f.Bold.HasValue && f.Bold.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Bold))
            {
                await this.WriteCssItemAsync("font-weight:bolder;", this._settings.Minify);
            }

            if (f.Italic.HasValue && f.Italic.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Italic))
            {
                await this.WriteCssItemAsync("font-style:italic;", this._settings.Minify);
            }

            if (f.Strike.HasValue && f.Strike.Value && EnumUtil.HasNotFlag(flags, eFontExclude.Strike))
            {
                await this.WriteCssItemAsync("text-decoration:line-through solid;", this._settings.Minify);
            }

            if (f.Underline.HasValue && f.Underline != ExcelUnderLineType.None && EnumUtil.HasNotFlag(flags, eFontExclude.Underline))
            {
                await this.WriteCssItemAsync("text-decoration:underline ", this._settings.Minify);

                switch (f.Underline.Value)
                {
                    case ExcelUnderLineType.Double:
                    case ExcelUnderLineType.DoubleAccounting:
                        await this.WriteCssItemAsync("double;", this._settings.Minify);

                        break;

                    default:
                        await this.WriteCssItemAsync("solid;", this._settings.Minify);

                        break;
                }
            }
        }

        private async Task WriteBorderStylesAsync(ExcelDxfBorderBase b)
        {
            if (b.HasValue)
            {
                eBorderExclude flags = this._settings.Css.Exclude.TableStyle.Border;

                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Top))
                {
                    await this.WriteBorderItemAsync(b.Top, "top");
                }

                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Bottom))
                {
                    await this.WriteBorderItemAsync(b.Bottom, "bottom");
                }

                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Left))
                {
                    await this.WriteBorderItemAsync(b.Left, "left");
                }

                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Right))
                {
                    await this.WriteBorderItemAsync(b.Right, "right");
                }
            }
        }

        private async Task WriteBorderStylesVerticalHorizontalAsync(ExcelDxfBorderBase b)
        {
            if (b.HasValue)
            {
                eBorderExclude flags = this._settings.Css.Exclude.TableStyle.Border;

                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Top))
                {
                    await this.WriteBorderItemAsync(b.Horizontal, "top");
                }

                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Bottom))
                {
                    await this.WriteBorderItemAsync(b.Horizontal, "bottom");
                }

                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Left))
                {
                    await this.WriteBorderItemAsync(b.Vertical, "left");
                }

                if (EnumUtil.HasNotFlag(flags, eBorderExclude.Right))
                {
                    await this.WriteBorderItemAsync(b.Vertical, "right");
                }
            }
        }

        private async Task WriteBorderItemAsync(ExcelDxfBorderItem bi, string suffix)
        {
            if (bi.HasValue && bi.Style != ExcelBorderStyle.None)
            {
                StringBuilder? sb = new StringBuilder();
                sb.Append(GetBorderItemLine(bi.Style.Value, suffix));

                if (bi.Color.HasValue)
                {
                    sb.Append($" {this.GetDxfColor(bi.Color)}");
                }

                sb.Append(";");

                await this.WriteCssItemAsync(sb.ToString(), this._settings.Minify);
            }
        }
    }
#endif
}