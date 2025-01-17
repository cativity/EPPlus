﻿using OfficeOpenXml.Utils.Extensions;
using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml;

/// <summary>
/// Represents a color in a vml.
/// </summary>
public class ExcelVmlDrawingColor : XmlHelper
{
    string _path;

    internal ExcelVmlDrawingColor(XmlNamespaceManager ns, XmlNode topNode, string path)
        : base(ns, topNode) =>
        this._path = path;

    /// <summary>
    /// A color string representing a color. Uses the HTML 4.0 color names, rgb decimal triplets or rgb hex triplets
    /// Example: 
    /// ColorString = "rgb(200,100, 0)"
    /// ColorString = "#FF0000"
    /// ColorString = "Red"
    /// ColorString = "#345" //This is the same as #334455
    /// </summary>
    public string ColorString
    {
        get => this.GetXmlNodeString(this._path);
        set => this.SetXmlNodeString(this._path, value);
    }

    /// <summary>
    /// Sets the Color string with the color supplied.
    /// </summary>
    /// <param name="color"></param>
    public void SetColor(Color color) => this.ColorString = "#" + (color.ToArgb() & 0xFFFFFF).ToString("X").PadLeft(6, '0');

    /// <summary>
    /// Gets the color for the color string
    /// </summary>
    /// <returns></returns>
    public Color GetColor() => GetColor(this.ColorString);

    internal static Color GetColor(string c)
    {
        if (string.IsNullOrEmpty(c))
        {
            return Color.Empty;
        }

        try
        {
            if (c.IndexOf("[", StringComparison.OrdinalIgnoreCase) > 0)
            {
                c = c.Substring(0, c.IndexOf("[")).Trim();
            }

            string? ts = c.Replace(" ", "");

            if (ts.StartsWith("rgb(", StringComparison.InvariantCultureIgnoreCase))
            {
                string[]? l = ts.Substring(4, ts.Length - 5).Split(',');

                if (l.Length == 3)
                {
                    return Color.FromArgb(0xFF, int.Parse(l[0]), int.Parse(l[1]), int.Parse(l[2]));
                }

                return Color.Empty;
            }
            else
            {
#if NETSTANDARD
                    return OfficeOpenXml.Compatibility.System.Drawing.ColorTranslator.FromHtml(c);
#else
                return ColorTranslator.FromHtml(c);
#endif
            }
        }
        catch
        {
            return Color.Empty;
        }
    }
}