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

using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style.XmlAccess;
using System.IO;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System;
using System.Collections.Generic;
using OfficeOpenXml.Utils;
using System.Text;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Interfaces;
using System.Linq;
using OfficeOpenXml.Export.HtmlExport.Exporters;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
#if !NET35 && !NET40
    internal partial class EpplusCssWriter : HtmlWriterBase
    {
        internal async Task RenderAdditionalAndFontCssAsync(string tableClass)
        {
            if (this._cssSettings.IncludeSharedClasses == false)
            {
                return;
            }

            await this.WriteClassAsync($"table.{tableClass}{{", this._settings.Minify);

            if (this._cssSettings.IncludeNormalFont)
            {
                ExcelNamedStyleXml? ns = this._ranges.First().Worksheet.Workbook.Styles.GetNormalStyle();

                if (ns != null)
                {
                    await this.WriteCssItemAsync($"font-family:{ns.Style.Font.Name};", this._settings.Minify);
                    await this.WriteCssItemAsync($"font-size:{ns.Style.Font.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", this._settings.Minify);
                }
            }

            foreach (KeyValuePair<string, string> item in this._cssSettings.AdditionalCssElements)
            {
                await this.WriteCssItemAsync($"{item.Key}:{item.Value};", this._settings.Minify);
            }

            await this.WriteClassEndAsync(this._settings.Minify);

            //Class for hidden rows.
            await this.WriteClassAsync($".{this._settings.StyleClassPrefix}hidden {{", this._settings.Minify);
            await this.WriteCssItemAsync($"display:none;", this._settings.Minify);
            await this.WriteClassEndAsync(this._settings.Minify);

            await this.WriteClassAsync($".{this._settings.StyleClassPrefix}al {{", this._settings.Minify);
            await this.WriteCssItemAsync($"text-align:left;", this._settings.Minify);
            await this.WriteClassEndAsync(this._settings.Minify);
            await this.WriteClassAsync($".{this._settings.StyleClassPrefix}ar {{", this._settings.Minify);
            await this.WriteCssItemAsync($"text-align:right;", this._settings.Minify);
            await this.WriteClassEndAsync(this._settings.Minify);

            List<ExcelWorksheet>? worksheets = this._ranges.Select(x => x.Worksheet).Distinct().ToList();

            foreach (ExcelWorksheet? ws in worksheets)
            {
                string? clsName = HtmlExportTableUtil.GetWorksheetClassName(this._settings.StyleClassPrefix, "dcw", ws, worksheets.Count > 1);
                await this.WriteClassAsync($".{clsName} {{", this._settings.Minify);

                await this.WriteCssItemAsync($"width:{ExcelColumn.ColumnWidthToPixels(Convert.ToDecimal(ws.DefaultColWidth), ws.Workbook.MaxFontWidth)}px;",
                                             this._settings.Minify);

                await this.WriteClassEndAsync(this._settings.Minify);

                clsName = HtmlExportTableUtil.GetWorksheetClassName(this._settings.StyleClassPrefix, "drh", ws, worksheets.Count > 1);
                await this.WriteClassAsync($".{clsName} {{", this._settings.Minify);
                await this.WriteCssItemAsync($"height:{(int)(ws.DefaultRowHeight / 0.75)}px;", this._settings.Minify);
                await this.WriteClassEndAsync(this._settings.Minify);
            }

            if (this._settings.Pictures.Include != ePictureInclude.Exclude && this._settings.Pictures.CssExclude.Alignment == false)
            {
                await this.WriteClassAsync($"td.{this._settings.StyleClassPrefix}image-cell {{", this._settings.Minify);

                if (this._settings.Pictures.AddMarginTop)
                {
                    await this.WriteCssItemAsync($"vertical-align:top;", this._settings.Minify);
                }
                else
                {
                    await this.WriteCssItemAsync($"vertical-align:middle;", this._settings.Minify);
                }

                if (this._settings.Pictures.AddMarginTop)
                {
                    await this.WriteCssItemAsync($"text-align:left;", this._settings.Minify);
                }
                else
                {
                    await this.WriteCssItemAsync($"text-align:center;", this._settings.Minify);
                }

                await this.WriteClassEndAsync(this._settings.Minify);
            }
        }

        internal async Task AddPictureToCssAsync(HtmlImage p)
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
                await this.WriteClassAsync($"img.{this._settings.StyleClassPrefix}image-{imageFileName}{{", this._settings.Minify);
                await this.WriteCssItemAsync($"content:url('data:{GetContentType(type.Value)};base64,{encodedImage}');", this._settings.Minify);

                if (this._settings.Pictures.Position != ePicturePosition.DontSet)
                {
                    await this.WriteCssItemAsync($"position:{this._settings.Pictures.Position.ToString().ToLower()};", this._settings.Minify);
                }

                if (p.FromColumnOff != 0 && this._settings.Pictures.AddMarginLeft)
                {
                    int leftOffset = p.FromColumnOff / ExcelDrawing.EMU_PER_PIXEL;
                    await this.WriteCssItemAsync($"margin-left:{leftOffset}px;", this._settings.Minify);
                }

                if (p.FromRowOff != 0 && this._settings.Pictures.AddMarginTop)
                {
                    int topOffset = p.FromRowOff / ExcelDrawing.EMU_PER_PIXEL;
                    await this.WriteCssItemAsync($"margin-top:{topOffset}px;", this._settings.Minify);
                }

                await this.WriteClassEndAsync(this._settings.Minify);
                _ = this._images.Add(pc.ImageHash);
            }

            await this.AddPicturePropertiesToCssAsync(p);
        }

        private async Task AddPicturePropertiesToCssAsync(HtmlImage image)
        {
            string imageName = HtmlExportTableUtil.GetClassName(image.Picture.Name, ((IPictureContainer)image.Picture).ImageHash);
            double width = image.Picture.GetPixelWidth();
            double height = image.Picture.GetPixelHeight();

            await this.WriteClassAsync($"img.{this._settings.StyleClassPrefix}image-prop-{imageName}{{", this._settings.Minify);

            if (this._settings.Pictures.KeepOriginalSize == false)
            {
                if (width != image.Picture.Image.Bounds.Width)
                {
                    await this.WriteCssItemAsync($"max-width:{width:F0}px;", this._settings.Minify);
                }

                if (height != image.Picture.Image.Bounds.Height)
                {
                    await this.WriteCssItemAsync($"max-height:{height:F0}px;", this._settings.Minify);
                }
            }

            if (image.Picture.Border.LineStyle != null && this._settings.Pictures.CssExclude.Border == false)
            {
                string? border = GetDrawingBorder(image.Picture);
                await this.WriteCssItemAsync($"border:{border};", this._settings.Minify);
            }

            await this.WriteClassEndAsync(this._settings.Minify);
        }

        internal async Task AddToCssAsync(ExcelStyles styles, int styleId, string styleClassPrefix, string cellStyleClassName)
        {
            ExcelXfs? xfs = styles.CellXfs[styleId];

            if (HasStyle(xfs))
            {
                if (this.IsAddedToCache(xfs, out int id) == false || this._addedToCss.Contains(id) == false)
                {
                    _ = this._addedToCss.Add(id);
                    await this.WriteClassAsync($".{styleClassPrefix}{cellStyleClassName}{id}{{", this._settings.Minify);

                    if (xfs.FillId > 0)
                    {
                        await this.WriteFillStylesAsync(xfs.Fill);
                    }

                    if (xfs.FontId > 0)
                    {
                        ExcelNamedStyleXml? ns = styles.GetNormalStyle();
                        await this.WriteFontStylesAsync(xfs.Font, ns.Style.Font);
                    }

                    if (xfs.BorderId > 0)
                    {
                        await this.WriteBorderStylesAsync(xfs.Border.Top, xfs.Border.Bottom, xfs.Border.Left, xfs.Border.Right);
                    }

                    await this.WriteStylesAsync(xfs);
                    await this.WriteClassEndAsync(this._settings.Minify);
                }
            }
        }

        internal async Task AddToCssAsync(ExcelStyles styles,
                                          int styleId,
                                          int bottomStyleId,
                                          int rightStyleId,
                                          string styleClassPrefix,
                                          string cellStyleClassName)
        {
            ExcelXfs? xfs = styles.CellXfs[styleId];
            ExcelXfs? bXfs = styles.CellXfs[bottomStyleId];
            ExcelXfs? rXfs = styles.CellXfs[rightStyleId];

            if (HasStyle(xfs) || bXfs.BorderId > 0 || rXfs.BorderId > 0)
            {
                if (this.IsAddedToCache(xfs, out int id, bottomStyleId, rightStyleId) == false || this._addedToCss.Contains(id) == false)
                {
                    _ = this._addedToCss.Add(id);
                    await this.WriteClassAsync($".{styleClassPrefix}{cellStyleClassName}{id}{{", this._settings.Minify);

                    if (xfs.FillId > 0)
                    {
                        this.WriteFillStyles(xfs.Fill);
                    }

                    if (xfs.FontId > 0)
                    {
                        ExcelNamedStyleXml? ns = styles.GetNormalStyle();
                        await this.WriteFontStylesAsync(xfs.Font, ns.Style.Font);
                    }

                    if (xfs.BorderId > 0 || bXfs.BorderId > 0 || rXfs.BorderId > 0)
                    {
                        await this.WriteBorderStylesAsync(xfs.Border.Top, bXfs.Border.Bottom, xfs.Border.Left, rXfs.Border.Right);
                    }

                    await this.WriteStylesAsync(xfs);
                    await this.WriteClassEndAsync(this._settings.Minify);
                }
            }
        }

        private async Task WriteStylesAsync(ExcelXfs xfs)
        {
            if (this._cssExclude.WrapText == false)
            {
                if (xfs.WrapText)
                {
                    await this.WriteCssItemAsync("white-space: break-spaces;", this._settings.Minify);
                }
                else
                {
                    await this.WriteCssItemAsync("white-space: nowrap;", this._settings.Minify);
                }
            }

            if (xfs.HorizontalAlignment != ExcelHorizontalAlignment.General && this._cssExclude.HorizontalAlignment == false)
            {
                string? hAlign = GetHorizontalAlignment(xfs);
                await this.WriteCssItemAsync($"text-align:{hAlign};", this._settings.Minify);
            }

            if (xfs.VerticalAlignment != ExcelVerticalAlignment.Bottom && this._cssExclude.VerticalAlignment == false)
            {
                string? vAlign = GetVerticalAlignment(xfs);
                await this.WriteCssItemAsync($"vertical-align:{vAlign};", this._settings.Minify);
            }

            if (xfs.TextRotation != 0 && this._cssExclude.TextRotation == false)
            {
                if (xfs.TextRotation == 255)
                {
                    await this.WriteCssItemAsync($"writing-mode:vertical-lr;;", this._settings.Minify);
                    await this.WriteCssItemAsync($"text-orientation:upright;", this._settings.Minify);
                }
                else
                {
                    if (xfs.TextRotation > 90)
                    {
                        await this.WriteCssItemAsync($"transform:rotate({xfs.TextRotation - 90}deg);", this._settings.Minify);
                    }
                    else
                    {
                        await this.WriteCssItemAsync($"transform:rotate({360 - xfs.TextRotation}deg);", this._settings.Minify);
                    }
                }
            }

            if (xfs.Indent > 0 && this._cssExclude.Indent == false)
            {
                await this.WriteCssItemAsync($"padding-left:{xfs.Indent * this._cssSettings.IndentValue}{this._cssSettings.IndentUnit};",
                                             this._settings.Minify);
            }
        }

        private async Task WriteBorderStylesAsync(ExcelBorderItemXml top, ExcelBorderItemXml bottom, ExcelBorderItemXml left, ExcelBorderItemXml right)
        {
            if (EnumUtil.HasNotFlag(this._borderExclude, eBorderExclude.Top))
            {
                await this.WriteBorderItemAsync(top, "top");
            }

            if (EnumUtil.HasNotFlag(this._borderExclude, eBorderExclude.Bottom))
            {
                await this.WriteBorderItemAsync(bottom, "bottom");
            }

            if (EnumUtil.HasNotFlag(this._borderExclude, eBorderExclude.Left))
            {
                await this.WriteBorderItemAsync(left, "left");
            }

            if (EnumUtil.HasNotFlag(this._borderExclude, eBorderExclude.Right))
            {
                await this.WriteBorderItemAsync(right, "right");
            }

            //TODO add Diagonal
            //WriteBorderItem(b.DiagonalDown, "right");
            //WriteBorderItem(b.DiagonalUp, "right");
        }

        private async Task WriteBorderItemAsync(ExcelBorderItemXml bi, string suffix)
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

                await this.WriteCssItemAsync(sb.ToString(), this._settings.Minify);
            }
        }

        private async Task WriteFontStylesAsync(ExcelFontXml f, ExcelFont nf)
        {
            if (string.IsNullOrEmpty(f.Name) == false && EnumUtil.HasNotFlag(this._fontExclude, eFontExclude.Name) && f.Name.Equals(nf.Name) == false)
            {
                await this.WriteCssItemAsync($"font-family:{f.Name};", this._settings.Minify);
            }

            if (f.Size > 0 && EnumUtil.HasNotFlag(this._fontExclude, eFontExclude.Size) && f.Size != nf.Size)
            {
                await this.WriteCssItemAsync($"font-size:{f.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", this._settings.Minify);
            }

            if (f.Color != null && f.Color.Exists && EnumUtil.HasNotFlag(this._fontExclude, eFontExclude.Color) && AreColorEqual(f.Color, nf.Color) == false)
            {
                await this.WriteCssItemAsync($"color:{this.GetColor(f.Color)};", this._settings.Minify);
            }

            if (f.Bold && EnumUtil.HasNotFlag(this._fontExclude, eFontExclude.Bold) && nf.Bold != f.Bold)
            {
                await this.WriteCssItemAsync("font-weight:bolder;", this._settings.Minify);
            }

            if (f.Italic && EnumUtil.HasNotFlag(this._fontExclude, eFontExclude.Italic) && nf.Italic != f.Italic)
            {
                await this.WriteCssItemAsync("font-style:italic;", this._settings.Minify);
            }

            if (f.Strike && EnumUtil.HasNotFlag(this._fontExclude, eFontExclude.Strike) && nf.Strike != f.Strike)
            {
                await this.WriteCssItemAsync("text-decoration:line-through solid;", this._settings.Minify);
            }

            if (f.UnderLineType != ExcelUnderLineType.None
                && EnumUtil.HasNotFlag(this._fontExclude, eFontExclude.Underline)
                && f.UnderLineType != nf.UnderLineType)
            {
                switch (f.UnderLineType)
                {
                    case ExcelUnderLineType.Double:
                    case ExcelUnderLineType.DoubleAccounting:
                        await this.WriteCssItemAsync("text-decoration:underline double;", this._settings.Minify);

                        break;

                    default:
                        await this.WriteCssItemAsync("text-decoration:underline solid;", this._settings.Minify);

                        break;
                }
            }
        }

        private async Task WriteFillStylesAsync(ExcelFillXml f)
        {
            if (this._cssExclude.Fill)
            {
                return;
            }

            if (f is ExcelGradientFillXml gf && gf.Type != ExcelFillGradientType.None)
            {
                await this.WriteGradientAsync(gf);
            }
            else
            {
                if (f.PatternType == ExcelFillStyle.Solid)
                {
                    await this.WriteCssItemAsync($"background-color:{this.GetColor(f.BackgroundColor)};", this._settings.Minify);
                }
                else
                {
                    await
                        this.WriteCssItemAsync($"{PatternFills.GetPatternSvg(f.PatternType, this.GetColor(f.BackgroundColor), this.GetColor(f.PatternColor))}",
                                               this._settings.Minify);
                }
            }
        }

        private async Task WriteGradientAsync(ExcelGradientFillXml gradient)
        {
            if (gradient.Type == ExcelFillGradientType.Linear)
            {
                await this._writer.WriteAsync($"background: linear-gradient({(gradient.Degree + 90) % 360}deg");
            }
            else
            {
                await this._writer.WriteAsync($"background:radial-gradient(ellipse {gradient.Right * 100}% {gradient.Bottom * 100}%");
            }

            await this._writer.WriteAsync($",{this.GetColor(gradient.GradientColor1)} 0%");
            await this._writer.WriteAsync($",{this.GetColor(gradient.GradientColor2)} 100%");

            await this._writer.WriteAsync(");");
        }

        public async Task FlushStreamAsync()
        {
            await this._writer.FlushAsync();
        }
    }
#endif
}