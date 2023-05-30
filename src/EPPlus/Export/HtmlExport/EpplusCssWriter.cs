/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/17/2021         EPPlus Software AB       Added Html Export
 *************************************************************************************************/

using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style.XmlAccess;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System;
using OfficeOpenXml.Utils;
using System.Text;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Interfaces;
using System.Linq;
using OfficeOpenXml.Export.HtmlExport.Exporters;

namespace OfficeOpenXml.Export.HtmlExport;

internal partial class EpplusCssWriter : HtmlWriterBase
{
    protected HtmlExportSettings _settings;
    protected CssExportSettings _cssSettings;
    protected CssExclude _cssExclude;
    List<ExcelRangeBase> _ranges;
    ExcelWorkbook _wb;
    ExcelTheme _theme;
    internal eFontExclude _fontExclude;
    internal eBorderExclude _borderExclude;
    internal HashSet<int> _addedToCss = new HashSet<int>();

    internal EpplusCssWriter(StreamWriter writer,
                             List<ExcelRangeBase> ranges,
                             HtmlExportSettings settings,
                             CssExportSettings cssSettings,
                             CssExclude cssExclude,
                             Dictionary<string, int> styleCache)
        : base(writer, styleCache)
    {
        this._settings = settings;
        this._cssSettings = cssSettings;
        this._cssExclude = cssExclude;
        this.Init(ranges);
    }

    internal EpplusCssWriter(Stream stream,
                             List<ExcelRangeBase> ranges,
                             HtmlExportSettings settings,
                             CssExportSettings cssSettings,
                             CssExclude cssExclude,
                             Dictionary<string, int> styleCache)
        : base(stream, settings.Encoding, styleCache)
    {
        this._settings = settings;
        this._cssSettings = cssSettings;
        this._cssExclude = cssExclude;
        this.Init(ranges);
    }

    private void Init(List<ExcelRangeBase> ranges)
    {
        this._ranges = ranges;
        this._wb = this._ranges[0].Worksheet.Workbook;

        if (this._wb.ThemeManager.CurrentTheme == null)
        {
            this._wb.ThemeManager.CreateDefaultTheme();
        }

        this._theme = this._wb.ThemeManager.CurrentTheme;
        this._borderExclude = this._cssExclude.Border;
        this._fontExclude = this._cssExclude.Font;
    }

    internal void RenderAdditionalAndFontCss(string tableClass)
    {
        if (this._cssSettings.IncludeSharedClasses == false)
        {
            return;
        }

        this.WriteClass($"table.{tableClass}{{", this._settings.Minify);

        if (this._cssSettings.IncludeNormalFont)
        {
            ExcelNamedStyleXml? ns = this._wb.Styles.GetNormalStyle();

            if (ns != null)
            {
                this.WriteCssItem($"font-family:{ns.Style.Font.Name};", this._settings.Minify);
                this.WriteCssItem($"font-size:{ns.Style.Font.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", this._settings.Minify);
            }
        }

        foreach (KeyValuePair<string, string> item in this._cssSettings.AdditionalCssElements)
        {
            this.WriteCssItem($"{item.Key}:{item.Value};", this._settings.Minify);
        }

        this.WriteClassEnd(this._settings.Minify);

        //Class for hidden rows.
        this.WriteClass($".{this._settings.StyleClassPrefix}hidden {{", this._settings.Minify);
        this.WriteCssItem($"display:none;", this._settings.Minify);
        this.WriteClassEnd(this._settings.Minify);

        this.WriteClass($".{this._settings.StyleClassPrefix}al {{", this._settings.Minify);
        this.WriteCssItem($"text-align:left;", this._settings.Minify);
        this.WriteClassEnd(this._settings.Minify);

        this.WriteClass($".{this._settings.StyleClassPrefix}ar {{", this._settings.Minify);
        this.WriteCssItem($"text-align:right;", this._settings.Minify);
        this.WriteClassEnd(this._settings.Minify);

        List<ExcelWorksheet>? worksheets = this._ranges.Select(x => x.Worksheet).Distinct().ToList();

        foreach (ExcelWorksheet? ws in worksheets)
        {
            string? clsName = HtmlExportTableUtil.GetWorksheetClassName(this._settings.StyleClassPrefix, "dcw", ws, worksheets.Count > 1);
            this.WriteClass($".{clsName} {{", this._settings.Minify);

            this.WriteCssItem($"width:{ExcelColumn.ColumnWidthToPixels(Convert.ToDecimal(ws.DefaultColWidth), ws.Workbook.MaxFontWidth)}px;",
                              this._settings.Minify);

            this.WriteClassEnd(this._settings.Minify);

            clsName = HtmlExportTableUtil.GetWorksheetClassName(this._settings.StyleClassPrefix, "drh", ws, worksheets.Count > 1);
            this.WriteClass($".{clsName} {{", this._settings.Minify);
            this.WriteCssItem($"height:{(int)(ws.DefaultRowHeight / 0.75)}px;", this._settings.Minify);
            this.WriteClassEnd(this._settings.Minify);
        }

        //Image alignment class
        if (this._settings.Pictures.Include != ePictureInclude.Exclude && this._settings.Pictures.CssExclude.Alignment == false)
        {
            this.WriteClass($"td.{this._settings.StyleClassPrefix}image-cell {{", this._settings.Minify);

            if (this._settings.Pictures.AddMarginTop)
            {
                this.WriteCssItem($"vertical-align:top;", this._settings.Minify);
            }
            else
            {
                this.WriteCssItem($"vertical-align:middle;", this._settings.Minify);
            }

            if (this._settings.Pictures.AddMarginTop)
            {
                this.WriteCssItem($"text-align:left;", this._settings.Minify);
            }
            else
            {
                this.WriteCssItem($"text-align:center;", this._settings.Minify);
            }

            this.WriteClassEnd(this._settings.Minify);
        }
    }

    internal void AddPictureToCss(HtmlImage p)
    {
        ExcelImage? img = p.Picture.Image;
        string encodedImage;
        ePictureType? type;

        if (img.Type == ePictureType.Emz || img.Type == ePictureType.Wmz)
        {
            encodedImage = Convert.ToBase64String(ImageReader.ExtractImage(img.ImageBytes, out type));
        }
        else
        {
            encodedImage = Convert.ToBase64String(img.ImageBytes);
            type = img.Type.Value;
        }

        if (type == null)
        {
            return;
        }

        IPictureContainer? pc = (IPictureContainer)p.Picture;

        if (this._images.Contains(pc.ImageHash) == false)
        {
            string imageFileName = HtmlExportImageUtil.GetPictureName(p);
            this.WriteClass($"img.{this._settings.StyleClassPrefix}image-{imageFileName}{{", this._settings.Minify);
            this.WriteCssItem($"content:url('data:{GetContentType(type.Value)};base64,{encodedImage}');", this._settings.Minify);

            if (this._settings.Pictures.Position != ePicturePosition.DontSet)
            {
                this.WriteCssItem($"position:{this._settings.Pictures.Position.ToString().ToLower()};", this._settings.Minify);
            }

            if (p.FromColumnOff != 0 && this._settings.Pictures.AddMarginLeft)
            {
                int leftOffset = p.FromColumnOff / ExcelDrawing.EMU_PER_PIXEL;
                this.WriteCssItem($"margin-left:{leftOffset}px;", this._settings.Minify);
            }

            if (p.FromRowOff != 0 && this._settings.Pictures.AddMarginTop)
            {
                int topOffset = p.FromRowOff / ExcelDrawing.EMU_PER_PIXEL;
                this.WriteCssItem($"margin-top:{topOffset}px;", this._settings.Minify);
            }

            this.WriteClassEnd(this._settings.Minify);
            _ = this._images.Add(pc.ImageHash);
        }

        this.AddPicturePropertiesToCss(p);
    }

    private void AddPicturePropertiesToCss(HtmlImage image)
    {
        string imageName = HtmlExportTableUtil.GetClassName(image.Picture.Name, ((IPictureContainer)image.Picture).ImageHash);
        double width = image.Picture.GetPixelWidth();
        double height = image.Picture.GetPixelHeight();

        this.WriteClass($"img.{this._settings.StyleClassPrefix}image-prop-{imageName}{{", this._settings.Minify);

        if (this._settings.Pictures.KeepOriginalSize == false)
        {
            if (width != image.Picture.Image.Bounds.Width)
            {
                this.WriteCssItem($"max-width:{width:F0}px;", this._settings.Minify);
            }

            if (height != image.Picture.Image.Bounds.Height)
            {
                this.WriteCssItem($"max-height:{height:F0}px;", this._settings.Minify);
            }
        }

        if (image.Picture.Border.LineStyle != null && this._settings.Pictures.CssExclude.Border == false)
        {
            string? border = GetDrawingBorder(image.Picture);
            this.WriteCssItem($"border:{border};", this._settings.Minify);
        }

        this.WriteClassEnd(this._settings.Minify);
    }

    private static string GetDrawingBorder(ExcelPicture picture)
    {
        Color color = picture.Border.Fill.Color;

        if (color.IsEmpty)
        {
            return "";
        }

        string lineStyle = $"{picture.Border.Width}px";

        switch (picture.Border.LineStyle.Value)
        {
            case eLineStyle.Solid:
                lineStyle += " solid";

                break;

            case eLineStyle.Dash:
            case eLineStyle.LongDashDot:
            case eLineStyle.LongDashDotDot:
            case eLineStyle.SystemDash:
            case eLineStyle.SystemDashDot:
            case eLineStyle.SystemDashDotDot:
                lineStyle += $" dashed";

                break;

            case eLineStyle.Dot:
                lineStyle += $" dot";

                break;
        }

        lineStyle += " #" + color.ToArgb().ToString("x8").Substring(2);

        return lineStyle;
    }

    private static object GetContentType(ePictureType type)
    {
        switch (type)
        {
            case ePictureType.Ico:
                return "image/vnd.microsoft.icon";

            case ePictureType.Jpg:
                return "image/jpeg";

            case ePictureType.Svg:
                return "image/svg+xml";

            case ePictureType.Tif:
                return "image/tiff";

            default:
                return $"image/{type}";
        }
    }

    internal void AddToCss(ExcelStyles styles, int styleId, string styleClassPrefix, string cellStyleClassName)
    {
        ExcelXfs? xfs = styles.CellXfs[styleId];

        if (HasStyle(xfs))
        {
            if (this.IsAddedToCache(xfs, out int id) == false || this._addedToCss.Contains(id) == false)
            {
                _ = this._addedToCss.Add(id);
                this.WriteClass($".{styleClassPrefix}{cellStyleClassName}{id}{{", this._settings.Minify);

                if (xfs.FillId > 0)
                {
                    this.WriteFillStyles(xfs.Fill);
                }

                if (xfs.FontId > 0)
                {
                    ExcelNamedStyleXml? ns = styles.GetNormalStyle();
                    this.WriteFontStyles(xfs.Font, ns.Style.Font);
                }

                if (xfs.BorderId > 0)
                {
                    this.WriteBorderStyles(xfs.Border.Top, xfs.Border.Bottom, xfs.Border.Left, xfs.Border.Right);
                }

                this.WriteStyles(xfs);
                this.WriteClassEnd(this._settings.Minify);
            }
        }
    }

    internal void AddToCss(ExcelStyles styles, int styleId, int bottomStyleId, int rightStyleId, string styleClassPrefix, string cellStyleClassName)
    {
        ExcelXfs? xfs = styles.CellXfs[styleId];
        ExcelXfs? bXfs = styles.CellXfs[bottomStyleId];
        ExcelXfs? rXfs = styles.CellXfs[rightStyleId];

        if (HasStyle(xfs) || bXfs.BorderId > 0 || rXfs.BorderId > 0)
        {
            if (this.IsAddedToCache(xfs, out int id, bottomStyleId, rightStyleId) == false || this._addedToCss.Contains(id) == false)
            {
                _ = this._addedToCss.Add(id);
                this.WriteClass($".{styleClassPrefix}{cellStyleClassName}{id}{{", this._settings.Minify);

                if (xfs.FillId > 0)
                {
                    this.WriteFillStyles(xfs.Fill);
                }

                if (xfs.FontId > 0)
                {
                    ExcelNamedStyleXml? ns = styles.GetNormalStyle();
                    this.WriteFontStyles(xfs.Font, ns.Style.Font);
                }

                if (xfs.BorderId > 0 || bXfs.BorderId > 0 || rXfs.BorderId > 0)
                {
                    this.WriteBorderStyles(xfs.Border.Top, bXfs.Border.Bottom, xfs.Border.Left, rXfs.Border.Right);
                }

                this.WriteStyles(xfs);
                this.WriteClassEnd(this._settings.Minify);
            }
        }
    }

    private bool IsAddedToCache(ExcelXfs xfs, out int id, int bottomStyleId = -1, int rightStyleId = -1)
    {
        string? key = GetStyleKey(xfs);

        if (bottomStyleId > -1)
        {
            key += bottomStyleId + "|" + rightStyleId;
        }

        if (this._styleCache.ContainsKey(key))
        {
            id = this._styleCache[key];

            return true;
        }
        else
        {
            id = this._styleCache.Count + 1;
            this._styleCache.Add(key, id);

            return false;
        }
    }

    private void WriteStyles(ExcelXfs xfs)
    {
        if (this._cssExclude.WrapText == false)
        {
            if (xfs.WrapText)
            {
                this.WriteCssItem("white-space: break-spaces;", this._settings.Minify);
            }
            else
            {
                this.WriteCssItem("white-space: nowrap;", this._settings.Minify);
            }
        }

        if (xfs.HorizontalAlignment != ExcelHorizontalAlignment.General && this._cssExclude.HorizontalAlignment == false)
        {
            string? hAlign = GetHorizontalAlignment(xfs);
            this.WriteCssItem($"text-align:{hAlign};", this._settings.Minify);
        }

        if (xfs.VerticalAlignment != ExcelVerticalAlignment.Bottom && this._cssExclude.VerticalAlignment == false)
        {
            string? vAlign = GetVerticalAlignment(xfs);
            this.WriteCssItem($"vertical-align:{vAlign};", this._settings.Minify);
        }

        if (xfs.TextRotation != 0 && this._cssExclude.TextRotation == false)
        {
            if (xfs.TextRotation == 255)
            {
                this.WriteCssItem($"writing-mode:vertical-lr;;", this._settings.Minify);
                this.WriteCssItem($"text-orientation:upright;", this._settings.Minify);
            }
            else
            {
                if (xfs.TextRotation > 90)
                {
                    this.WriteCssItem($"transform:rotate({xfs.TextRotation - 90}deg);", this._settings.Minify);
                }
                else
                {
                    this.WriteCssItem($"transform:rotate({360 - xfs.TextRotation}deg);", this._settings.Minify);
                }
            }
        }

        if (xfs.Indent > 0 && this._cssExclude.Indent == false)
        {
            this.WriteCssItem($"padding-left:{xfs.Indent * this._cssSettings.IndentValue}{this._cssSettings.IndentUnit};", this._settings.Minify);
        }
    }

    private void WriteBorderStyles(ExcelBorderItemXml top, ExcelBorderItemXml bottom, ExcelBorderItemXml left, ExcelBorderItemXml right)
    {
        if (EnumUtil.HasNotFlag(this._borderExclude, eBorderExclude.Top))
        {
            this.WriteBorderItem(top, "top");
        }

        if (EnumUtil.HasNotFlag(this._borderExclude, eBorderExclude.Bottom))
        {
            this.WriteBorderItem(bottom, "bottom");
        }

        if (EnumUtil.HasNotFlag(this._borderExclude, eBorderExclude.Left))
        {
            this.WriteBorderItem(left, "left");
        }

        if (EnumUtil.HasNotFlag(this._borderExclude, eBorderExclude.Right))
        {
            this.WriteBorderItem(right, "right");
        }

        //TODO add Diagonal
        //WriteBorderItem(b.DiagonalDown, "right");
        //WriteBorderItem(b.DiagonalUp, "right");
    }

    private void WriteBorderItem(ExcelBorderItemXml bi, string suffix)
    {
        if (bi.Style != ExcelBorderStyle.None)
        {
            StringBuilder? sb = new StringBuilder();
            _ = sb.Append(GetBorderItemLine(bi.Style, suffix));

            if (bi.Color != null && bi.Color.Exists)
            {
                _ = sb.Append($" {this.GetColor(bi.Color)}");
            }

            _ = sb.Append(";");

            this.WriteCssItem(sb.ToString(), this._settings.Minify);
        }
    }

    private void WriteFontStyles(ExcelFontXml f, ExcelFont nf)
    {
        if (string.IsNullOrEmpty(f.Name) == false && EnumUtil.HasNotFlag(this._fontExclude, eFontExclude.Name) && f.Name.Equals(nf.Name) == false)
        {
            this.WriteCssItem($"font-family:{f.Name};", this._settings.Minify);
        }

        if (f.Size > 0 && EnumUtil.HasNotFlag(this._fontExclude, eFontExclude.Size) && f.Size != nf.Size)
        {
            this.WriteCssItem($"font-size:{f.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", this._settings.Minify);
        }

        if (f.Color != null && f.Color.Exists && EnumUtil.HasNotFlag(this._fontExclude, eFontExclude.Color) && AreColorEqual(f.Color, nf.Color) == false)
        {
            this.WriteCssItem($"color:{this.GetColor(f.Color)};", this._settings.Minify);
        }

        if (f.Bold && EnumUtil.HasNotFlag(this._fontExclude, eFontExclude.Bold) && nf.Bold != f.Bold)
        {
            this.WriteCssItem("font-weight:bolder;", this._settings.Minify);
        }

        if (f.Italic && EnumUtil.HasNotFlag(this._fontExclude, eFontExclude.Italic) && nf.Italic != f.Italic)
        {
            this.WriteCssItem("font-style:italic;", this._settings.Minify);
        }

        if (f.Strike && EnumUtil.HasNotFlag(this._fontExclude, eFontExclude.Strike) && nf.Strike != f.Strike)
        {
            this.WriteCssItem("text-decoration:line-through solid;", this._settings.Minify);
        }

        if (f.UnderLineType != ExcelUnderLineType.None && EnumUtil.HasNotFlag(this._fontExclude, eFontExclude.Underline) && f.UnderLineType != nf.UnderLineType)
        {
            switch (f.UnderLineType)
            {
                case ExcelUnderLineType.Double:
                case ExcelUnderLineType.DoubleAccounting:
                    this.WriteCssItem("text-decoration:underline double;", this._settings.Minify);

                    break;

                default:
                    this.WriteCssItem("text-decoration:underline solid;", this._settings.Minify);

                    break;
            }
        }
    }

    private static bool AreColorEqual(ExcelColorXml c1, ExcelColor c2)
    {
        if (c1.Tint != c2.Tint)
        {
            return false;
        }

        if (c1.Indexed >= 0)
        {
            return c1.Indexed == c2.Indexed;
        }
        else if (string.IsNullOrEmpty(c1.Rgb) == false)
        {
            return c1.Rgb == c2.Rgb;
        }
        else if (c1.Theme != null)
        {
            return c1.Theme == c2.Theme;
        }
        else
        {
            return c1.Auto == c2.Auto;
        }
    }

    private void WriteFillStyles(ExcelFillXml f)
    {
        if (this._cssExclude.Fill)
        {
            return;
        }

        if (f is ExcelGradientFillXml gf && gf.Type != ExcelFillGradientType.None)
        {
            this.WriteGradient(gf);
        }
        else
        {
            if (f.PatternType == ExcelFillStyle.Solid)
            {
                this.WriteCssItem($"background-color:{this.GetColor(f.BackgroundColor)};", this._settings.Minify);
            }
            else
            {
                this.WriteCssItem($"{PatternFills.GetPatternSvg(f.PatternType, this.GetColor(f.BackgroundColor), this.GetColor(f.PatternColor))}",
                                  this._settings.Minify);
            }
        }
    }

    private void WriteGradient(ExcelGradientFillXml gradient)
    {
        if (gradient.Type == ExcelFillGradientType.Linear)
        {
            this._writer.Write($"background: linear-gradient({(gradient.Degree + 90) % 360}deg");
        }
        else
        {
            this._writer.Write($"background:radial-gradient(ellipse {gradient.Right * 100}% {gradient.Bottom * 100}%");
        }

        this._writer.Write($",{this.GetColor(gradient.GradientColor1)} 0%");
        this._writer.Write($",{this.GetColor(gradient.GradientColor2)} 100%");

        this._writer.Write(");");
    }

    private string GetColor(ExcelColorXml c)
    {
        Color ret;

        if (!string.IsNullOrEmpty(c.Rgb))
        {
            if (int.TryParse(c.Rgb, NumberStyles.HexNumber, null, out int hex))
            {
                ret = Color.FromArgb(hex);
            }
            else
            {
                ret = Color.Empty;
            }
        }
        else if (c.Theme.HasValue)
        {
            ret = Utils.ColorConverter.GetThemeColor(this._theme, c.Theme.Value);
        }
        else if (c.Indexed >= 0)
        {
            ret = ExcelColor.GetIndexedColor(c.Indexed);
        }
        else
        {
            //Automatic, set to black.
            ret = Color.Black;
        }

        if (c.Tint != 0)
        {
            ret = Utils.ColorConverter.ApplyTint(ret, Convert.ToDouble(c.Tint));
        }

        return "#" + ret.ToArgb().ToString("x8").Substring(2);
    }

    public void FlushStream() => this._writer.Flush();
}