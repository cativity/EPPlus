using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport;

internal static class HtmlRichText
{
    internal static void GetRichTextStyle(ExcelRichText rt, StringBuilder sb)
    {
        if (rt.Bold)
        {
            _ = sb.Append("font-weight:bolder;");
        }

        if (rt.Italic)
        {
            _ = sb.Append("font-style:italic;");
        }

        if (rt.UnderLine)
        {
            _ = sb.Append("text-decoration:underline solid;");
        }

        if (rt.Strike)
        {
            _ = sb.Append("text-decoration:line-through solid;");
        }

        if (rt.Size > 0)
        {
            _ = sb.Append($"font-size:{rt.Size.ToString("g", CultureInfo.InvariantCulture)}pt;");
        }

        if (string.IsNullOrEmpty(rt.FontName) == false)
        {
            _ = sb.Append($"font-family:{rt.FontName};");
        }

        if (rt.Color.IsEmpty == false)
        {
            _ = sb.Append("color:#" + rt.Color.ToArgb().ToString("x8").Substring(2));
        }
    }
}