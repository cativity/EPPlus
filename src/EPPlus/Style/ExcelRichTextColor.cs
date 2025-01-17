﻿/*************************************************************************************************
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

using OfficeOpenXml.Drawing;
using OfficeOpenXml.Packaging.Ionic;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.Utils.Extensions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.Globalization;

namespace OfficeOpenXml.Style;

public class ExcelRichTextColor : XmlHelper
{
    private ExcelRichText _rt;

    internal ExcelRichTextColor(XmlNamespaceManager ns, XmlNode topNode, ExcelRichText rt)
        : base(ns, topNode) =>
        this._rt = rt;

    /// <summary>
    /// Gets the rgb color depending in <see cref="Rgb"/>, <see cref="Theme"/> and <see cref="Tint"/>
    /// </summary>
    public Color Color => this._rt.Color;

    /// <summary>
    /// The rgb color value set in the file.
    /// </summary>
    public Color Rgb
    {
        get
        {
            string? col = this.GetXmlNodeString(ExcelRichText.COLOR_PATH);

            if (string.IsNullOrEmpty(col))
            {
                return Color.Empty;
            }

            return Color.FromArgb(int.Parse(col, NumberStyles.AllowHexSpecifier));
        }
        set
        {
            this._rt._collection.ConvertRichtext();

            if (value == Color.Empty)
            {
                this.DeleteNode(ExcelRichText.COLOR_PATH);
            }
            else
            {
                this.SetXmlNodeString(ExcelRichText.COLOR_PATH, value.ToArgb().ToString("X"));
            }

            if (this._rt._callback != null)
            {
                this._rt._callback();
            }
        }
    }

    /// <summary>
    /// The color theme.
    /// </summary>
    public eThemeSchemeColor? Theme
    {
        get => this.GetXmlNodeString(ExcelRichText.COLOR_THEME_PATH).ToEnum<eThemeSchemeColor>();
        set
        {
            this._rt._collection.ConvertRichtext();
            string? v = value.ToEnumString();

            if (v == null)
            {
                this.DeleteNode(ExcelRichText.COLOR_THEME_PATH);
            }
            else
            {
                this.SetXmlNodeString(ExcelRichText.COLOR_THEME_PATH, v);
            }

            if (this._rt._callback != null)
            {
                this._rt._callback();
            }
        }
    }

    /// <summary>
    /// The tint value for the color.
    /// </summary>
    public double? Tint
    {
        get => this.GetXmlNodeDoubleNull(ExcelRichText.COLOR_TINT_PATH);
        set
        {
            this._rt._collection.ConvertRichtext();
            this.SetXmlNodeDouble(ExcelRichText.COLOR_TINT_PATH, value, true);

            if (this._rt._callback != null)
            {
                this._rt._callback();
            }
        }
    }
}