using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Utils;
using System;
using System.IO;
using System.Text;
using OfficeOpenXml.Core.CellStore;
#if !NET35 && !NET40
using System.Threading.Tasks;
namespace OfficeOpenXml
{
    internal abstract partial class JsonExport
    {
        internal protected async Task WriteCellDataAsync(StreamWriter sw, ExcelRangeBase dr, int headerRows)
        {
            bool dtOnCell = this._settings.AddDataTypesOn == eDataTypeOn.OnCell;
            ExcelWorksheet ws = dr.Worksheet;
            Uri uri = null;
            int commentIx = 0;
            await this.WriteItemAsync(sw, $"\"{this._settings.RowsElementName}\":[", true);
            int fromRow = dr._fromRow + headerRows;
            for (int r = fromRow; r <= dr._toRow; r++)
            {
                await this.WriteStartAsync(sw);
                await this.WriteItemAsync(sw, $"\"{this._settings.CellsElementName}\":[", true);
                for (int c = dr._fromCol; c <= dr._toCol; c++)
                {
                    ExcelValue cv = ws.GetCoreValueInner(r, c);
                    string? t = JsonEscape(ValueToTextHandler.GetFormattedText(cv._value, ws.Workbook, cv._styleId, false, this._settings.Culture));
                    await this.WriteStartAsync(sw);
                    bool hasHyperlink = this._settings.WriteHyperlinks && ws._hyperLinks.Exists(r, c, ref uri);
                    bool hasComment = this._settings.WriteComments && ws._commentsStore.Exists(r, c, ref commentIx);
                    if (cv._value == null)
                    {
                        await this.WriteItemAsync(sw, $"\"t\":\"{t}\"");
                    }
                    else
                    {
                        string? v = JsonEscape(HtmlRawDataProvider.GetRawValue(cv._value));
                        await this.WriteItemAsync(sw, $"\"v\":\"{v}\",");
                        await this.WriteItemAsync(sw, $"\"t\":\"{t}\"", false, dtOnCell || hasHyperlink || hasComment);
                        if (dtOnCell)
                        {
                            string? dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(cv._value);
                            await this.WriteItemAsync(sw, $"\"dt\":\"{dt}\"", false, hasHyperlink || hasComment);
                        }
                    }

                    if (hasHyperlink)
                    {
                        await this.WriteItemAsync(sw, $"\"uri\":\"{JsonEscape(uri?.OriginalString)}\"", false, hasComment);
                    }

                    if (hasComment)
                    {
                        ExcelComment? comment = ws.Comments[commentIx];
                        await this.WriteItemAsync(sw, $"\"comment\":\"{comment.Text}\"");
                    }

                    if (c == dr._toCol)
                    {
                        await this.WriteEndAsync(sw, "}");
                    }
                    else
                    {
                        await this.WriteEndAsync(sw, "},");
                    }
                }
                await this.WriteEndAsync(sw, "]");
                if (r == dr._toRow)
                {
                    await this.WriteEndAsync(sw);
                }
                else
                {
                    await this.WriteEndAsync(sw, "},");
                }
            }
            await this.WriteEndAsync(sw, "]");
            await this.WriteEndAsync(sw);
        }
        internal protected async Task WriteItemAsync(StreamWriter sw, string v, bool indent = false, bool addComma = false)
        {
            if (addComma)
            {
                v += ",";
            }

            if (this._minify)
            {
                await sw.WriteAsync(v);
            }
            else
            {
                await sw.WriteLineAsync(this._indent + v);
                if (indent)
                {
                    this._indent += "  ";
                }
            }
        }

        internal protected async Task WriteStartAsync(StreamWriter sw)
        {
            if (this._minify)
            {
                await sw.WriteAsync("{");
            }
            else
            {
                await sw.WriteLineAsync($"{this._indent}{{");
                this._indent += "  ";
            }
        }
        internal protected async Task WriteEndAsync(StreamWriter sw, string bracket = "}")
        {
            if (this._minify)
            {
                await sw.WriteAsync(bracket);
            }
            else
            {
                this._indent = this._indent.Substring(0, this._indent.Length - 2);
                await sw.WriteLineAsync($"{this._indent}{bracket}");
            }
        }
    }
}
#endif
