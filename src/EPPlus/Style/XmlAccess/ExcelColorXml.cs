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
using System.Xml;
using System.Globalization;
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.Style.XmlAccess;

/// <summary>
/// Xml access class for color
/// </summary>
public sealed class ExcelColorXml : StyleXmlHelper
{
    internal ExcelColorXml(XmlNamespaceManager nameSpaceManager)
        : base(nameSpaceManager)
    {
        this._auto = false;
        this._theme = null;
        this._tint = 0;
        this._rgb = "";
        this._indexed = int.MinValue;
    }

    internal ExcelColorXml(XmlNamespaceManager nsm, XmlNode topNode)
        : base(nsm, topNode)
    {
        if (topNode == null)
        {
            this.Exists = false;
        }
        else
        {
            this.Exists = true;
            this._auto = this.GetXmlNodeBool("@auto");
            int? v = this.GetXmlNodeIntNull("@theme");

            if (v.HasValue && v >= 0 && v <= 11)
            {
                this._theme = (eThemeSchemeColor)v;
            }

            this._tint = this.GetXmlNodeDecimalNull("@tint") ?? decimal.MinValue;
            this._rgb = this.GetXmlNodeString("@rgb");
            this._indexed = this.GetXmlNodeIntNull("@indexed") ?? int.MinValue;
        }
    }

    internal override string Id
    {
        get { return this._auto.ToString() + "|" + this._theme?.ToString() + "|" + this._tint + "|" + this._rgb + "|" + this._indexed; }
    }

    bool _auto;

    /// <summary>
    /// Set the color to automatic
    /// </summary>
    public bool Auto
    {
        get { return this._auto; }
        set
        {
            this.Clear();
            this._auto = value;
            this.Exists = true;
        }
    }

    eThemeSchemeColor? _theme;

    /// <summary>
    /// Theme color value
    /// </summary>
    public eThemeSchemeColor? Theme
    {
        get { return this._theme; }
        set
        {
            this.Clear();
            this._theme = value;
            this.Exists = true;
        }
    }

    decimal _tint;

    /// <summary>
    /// The Tint value for the color
    /// </summary>
    public decimal Tint
    {
        get
        {
            if (this._tint == decimal.MinValue)
            {
                return 0;
            }
            else
            {
                return this._tint;
            }
        }
        set
        {
            this._tint = value;
            this.Exists = true;
        }
    }

    string _rgb;

    /// <summary>
    /// The RGB value
    /// </summary>
    public string Rgb
    {
        get { return this._rgb; }
        set
        {
            this._rgb = value;
            this.Exists = true;
            this._indexed = int.MinValue;
            this._auto = false;
        }
    }

    int _indexed;

    /// <summary>
    /// Indexed color value.
    /// Returns int.MinValue if indexed colors are not used.
    /// </summary>
    public int Indexed
    {
        get { return this._indexed; }
        set
        {
            if (value < 0 || value > 65)
            {
                throw new ArgumentOutOfRangeException("Index out of range");
            }

            this.Clear();
            this._indexed = value;
            this.Exists = true;
        }
    }

    internal void Clear()
    {
        this._theme = null;
        this._tint = decimal.MinValue;
        this._indexed = int.MinValue;
        this._rgb = "";
        this._auto = false;
    }

    /// <summary>
    /// Sets the color
    /// </summary>
    /// <param name="color">The color</param>
    public void SetColor(System.Drawing.Color color)
    {
        this.Clear();
        this._rgb = color.ToArgb().ToString("X");
    }

    /// <summary>
    /// Sets a theme color
    /// </summary>
    /// <param name="themeColorType">The theme color</param>
    public void SetColor(eThemeSchemeColor themeColorType)
    {
        this.Clear();
        this._theme = themeColorType;
    }

    /// <summary>
    /// Sets an indexed color
    /// </summary>
    /// <param name="indexedColor">The indexed color</param>
    public void SetColor(ExcelIndexedColor indexedColor)
    {
        this.Clear();
        this._indexed = (int)indexedColor;
    }

    internal ExcelColorXml Copy()
    {
        return new ExcelColorXml(this.NameSpaceManager)
        {
            _indexed = this._indexed, _tint = this._tint, _rgb = this._rgb, _theme = this._theme, _auto = this._auto, Exists = this.Exists
        };
    }

    internal override XmlNode CreateXmlNode(XmlNode topNode)
    {
        this.TopNode = topNode;

        if (this._rgb != "")
        {
            this.SetXmlNodeString("@rgb", this._rgb);
        }
        else if (this._indexed >= 0)
        {
            this.SetXmlNodeString("@indexed", this._indexed.ToString());
        }
        else if (this.Theme.HasValue)
        {
            this.SetXmlNodeString("@theme", ((int)this._theme).ToString(CultureInfo.InvariantCulture));
        }
        else
        {
            this.SetXmlNodeBool("@auto", this._auto);
        }

        if (this._tint != decimal.MinValue)
        {
            this.SetXmlNodeString("@tint", this._tint.ToString(CultureInfo.InvariantCulture));
        }

        return this.TopNode;
    }

    /// <summary>
    /// True if the record exists in the underlaying xml
    /// </summary>
    internal bool Exists { get; private set; }
}