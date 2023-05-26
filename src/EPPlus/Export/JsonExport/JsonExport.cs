using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Utils;
using System;
using System.IO;
using System.Text;
using OfficeOpenXml.Core.CellStore;

namespace OfficeOpenXml;

internal abstract partial class JsonExport
{
    private JsonExportSettings _settings;
    protected string _indent = "";
    protected bool _minify;
    internal JsonExport(JsonExportSettings settings)
    {
        this._settings = settings;
        this._minify = settings.Minify;
    }
    internal protected void WriteCellData(StreamWriter sw, ExcelRangeBase dr, int headerRows)
    {
        bool dtOnCell = this._settings.AddDataTypesOn == eDataTypeOn.OnCell;
        ExcelWorksheet ws = dr.Worksheet;
        Uri uri = null;
        int commentIx = 0;
        this.WriteItem(sw, $"\"{this._settings.RowsElementName}\":[", true);
        int fromRow = dr._fromRow + headerRows;
        for (int r = fromRow; r <= dr._toRow; r++)
        {
            this.WriteStart(sw);
            this.WriteItem(sw, $"\"{this._settings.CellsElementName}\":[", true);
            for (int c = dr._fromCol; c <= dr._toCol; c++)
            {
                ExcelValue cv = ws.GetCoreValueInner(r, c);
                string? t = JsonEscape(ValueToTextHandler.GetFormattedText(cv._value, ws.Workbook, cv._styleId, false, this._settings.Culture));
                this.WriteStart(sw);
                bool hasHyperlink = this._settings.WriteHyperlinks && ws._hyperLinks.Exists(r, c, ref uri);
                bool hasComment = this._settings.WriteComments && ws._commentsStore.Exists(r, c, ref commentIx);
                if (cv._value == null)
                {
                    this.WriteItem(sw, $"\"t\":\"{t}\"");
                }
                else
                {
                    string? v = JsonEscape(HtmlRawDataProvider.GetRawValue(cv._value));
                    this.WriteItem(sw, $"\"v\":\"{v}\",");
                    this.WriteItem(sw, $"\"t\":\"{t}\"", false, dtOnCell || hasHyperlink || hasComment);
                    if (dtOnCell)
                    {
                        string? dt = HtmlRawDataProvider.GetHtmlDataTypeFromValue(cv._value);
                        this.WriteItem(sw, $"\"dt\":\"{dt}\"", false, hasHyperlink  || hasComment);
                    }
                }

                if (hasHyperlink)
                {
                    this.WriteItem(sw, $"\"uri\":\"{JsonEscape(uri?.OriginalString)}\"", false, hasComment);
                }

                if(hasComment)
                {
                    ExcelComment? comment = ws.Comments[commentIx];
                    this.WriteItem(sw, $"\"comment\":\"{comment.Text}\"");
                }

                if(c == dr._toCol)
                {
                    this.WriteEnd(sw, "}");
                }
                else
                {
                    this.WriteEnd(sw,"},");
                }
            }

            this.WriteEnd(sw,"]");
            if (r == dr._toRow)
            {
                this.WriteEnd(sw);
            }
            else
            {
                this.WriteEnd(sw, "},");
            }
        }

        this.WriteEnd(sw, "]");
        this.WriteEnd(sw);
    }
    internal static string JsonEscape(string s)
    {
        if (s == null)
        {
            return "";
        }

        StringBuilder? sb = new StringBuilder();
        foreach (char c in s)
        {
            switch (c)
            {
                case '\\':
                    sb.Append("\\\\");
                    break;
                case '"':
                    sb.Append("\\\"");
                    break;
                case '\b':
                    sb.Append("\\b");
                    break;
                case '\f':
                    sb.Append("\\f");
                    break;
                case '\n':
                    sb.Append("\\n");
                    break;
                case '\r':
                    sb.Append("\\r");
                    break;
                case '\t':
                    sb.Append("\\t");
                    break;
                default:
                    if (c < 0x20)
                    {
                        sb.Append($"\\u{((short)c):X4}");
                    }
                    else
                    {
                        sb.Append(c);
                    }
                    break;
            }
        }
        return sb.ToString();
    }
    internal protected void WriteItem(StreamWriter sw, string v, bool indent=false, bool addComma=false)
    {
        if (addComma)
        {
            v += ",";
        }

        if (this._minify)
        {
            sw.Write(v);
        }
        else
        {
            sw.WriteLine(this._indent + v);
            if (indent)
            {
                this._indent += "  ";
            }
        }
    }

    internal protected void WriteStart(StreamWriter sw)
    {
        if (this._minify)
        {
            
            sw.Write("{");
        }
        else
        {
            sw.WriteLine($"{this._indent}{{");
            this._indent += "  ";
        }
    }
    internal protected void WriteEnd(StreamWriter sw, string bracket="}")
    {
        if (this._minify)
        {
            sw.Write(bracket);
        }
        else
        {
            this._indent = this._indent.Substring(0, this._indent.Length - 2);
            sw.WriteLine($"{this._indent}{bracket}");
        }
    }
}