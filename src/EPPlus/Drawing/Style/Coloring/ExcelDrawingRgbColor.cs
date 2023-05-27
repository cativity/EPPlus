/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Coloring;

/// <summary>
/// Represents a RGB color
/// </summary>
public class ExcelDrawingRgbColor : XmlHelper
{
    internal ExcelDrawingRgbColor(XmlNamespaceManager nsm, XmlNode topNode) : base (nsm, topNode)
    {
    }
    /// <summary>
    /// The color
    /// </summary>s
    public Color Color
    {
        get
        {
            string? s = this.GetXmlNodeString("@val");
            return GetColorFromString(s);
        }
        set
        {
            this.SetXmlNodeString("@val", (value.ToArgb() & 0xFFFFFF).ToString("X").PadLeft(6, '0'));
        }
    }
    internal static Color GetColorFromString(string s)
    {
        if (s.Length == 6)
        {
            s = "FF" + s;
        }

        if (int.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int n))
        {
            return Color.FromArgb(n);
        }
        else
        {
            return Color.Empty;
        }
    }

    internal const string NodeName = "a:srgbClr";
    internal static void SetXml(XmlNamespaceManager nsm, XmlNode node, bool doInit = false)
    {
            
    }
    internal static void GetXml()
    {
    }
    internal void GetHsl(out double hue, out double saturation, out double luminance)
    {
        GetHslColor(this.Color.R, this.Color.G, this.Color.B, out hue, out saturation, out luminance);
    }

    internal static void GetHslColor(Color c, out double hue, out double saturation, out double luminance)
    {
        GetHslColor(c.R, c.G, c.B, out hue, out saturation, out luminance);
    }
    internal static void GetHslColor(byte red, byte green, byte blue, out double hue, out double saturation, out double luminance)
    {
        //Created using formulas here...https://www.rapidtables.com/convert/color/rgb-to-hsl.html
        double r = red / 255D;
        double g = green / 255D;
        double b = blue / 255D;

        double[]? ix = new double[]{ r, g, b };
        double cMax = ix.Max();
        double cMin = ix.Min();
        double delta = cMax - cMin;


        if (delta == 0)
        {
            hue = 0;
        }
        else if (cMax == r)
        {
            hue = 60 * ((g - b) / delta % 6);
        }
        else if (cMax == g)
        {
            hue = 60 * (((b - r) / delta) + 2);
        }
        else
        {
            hue = 60 * (((r - g) / delta) + 4);
        }
           
        if (hue < 0)
        {
            hue += 360;
        }

        luminance = (cMax + cMin) / 2;
        saturation = delta == 0 ? 0 : delta / (1 - Math.Abs((2 * luminance) - 1));
    }
}