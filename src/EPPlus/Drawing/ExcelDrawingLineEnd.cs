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

namespace OfficeOpenXml.Drawing;

/// <summary>
/// Properties for drawing line ends
/// </summary>
public sealed class ExcelDrawingLineEnd : XmlHelper
{
    string _linePath;
    private readonly Action _init;

    internal ExcelDrawingLineEnd(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string linePath, Action init)
        : base(nameSpaceManager, topNode)
    {
        this._linePath = linePath;
        this._init = init;

        this.SchemaNodeOrder = new string[]
        {
            "noFill", "solidFill", "gradFill", "pattFill", "prstDash", "custDash", "round", "bevel", "miter", "headEnd", "tailEnd"
        };
    }

    string _stylePath = "/@type";

    /// <summary>
    /// The shapes line end decoration
    /// </summary>
    public eEndStyle? Style
    {
        get { return TranslateEndStyle(this.GetXmlNodeString(this._linePath + this._stylePath)); }
        set
        {
            this._init();

            if (value == null)
            {
                this.DeleteNode(this._linePath + this._stylePath);
            }
            else
            {
                this.SetXmlNodeString(this._linePath + this._stylePath, TranslateEndStyleText(value.Value));
            }
        }
    }

    string _widthPath = "/@w";

    /// <summary>
    /// The line start/end width in relation to the line width
    /// </summary>
    public eEndSize? Width
    {
        get { return TranslateEndSize(this.GetXmlNodeString(this._linePath + this._widthPath)); }
        set
        {
            this._init();

            if (value == null)
            {
                this.DeleteNode(this._linePath + this._widthPath);
            }
            else
            {
                this.SetXmlNodeString(this._linePath + this._widthPath, TranslateEndSizeText(value.Value));
            }
        }
    }

    string _heightPath = "/@len";

    /// <summary>
    /// The line start/end height in relation to the line height
    /// </summary>
    public eEndSize? Height
    {
        get { return TranslateEndSize(this.GetXmlNodeString(this._linePath + this._heightPath)); }
        set
        {
            this._init();

            if (value == null)
            {
                this.DeleteNode(this._linePath + this._heightPath);
            }
            else
            {
                this.SetXmlNodeString(this._linePath + this._heightPath, TranslateEndSizeText(value.Value));
            }
        }
    }

    #region "Translate Enum functions"

    private static string TranslateEndStyleText(eEndStyle value)
    {
        return value.ToString().ToLower();
    }

    private static eEndStyle? TranslateEndStyle(string text)
    {
        switch (text)
        {
            case "none":
            case "arrow":
            case "diamond":
            case "oval":
            case "stealth":
            case "triangle":
                return (eEndStyle)Enum.Parse(typeof(eEndStyle), text, true);

            default:
                return null;
        }
    }

    private string GetCreateLinePath(bool doCreate)
    {
        if (string.IsNullOrEmpty(this._linePath))
        {
            return "";
        }
        else
        {
            if (doCreate)
            {
                this.CreateNode(this._linePath, false);
            }

            return this._linePath + "/";
        }
    }

    private static string TranslateEndSizeText(eEndSize value)
    {
        string text = value.ToString();

        switch (value)
        {
            case eEndSize.Small:
                return "sm";

            case eEndSize.Medium:
                return "med";

            case eEndSize.Large:
                return "lg";

            default:
                return null;
        }
    }

    private static eEndSize? TranslateEndSize(string text)
    {
        switch (text)
        {
            case "sm":
                return eEndSize.Small;

            case "med":
                return eEndSize.Medium;

            case "lg":
                return eEndSize.Large;

            default:
                return null;
        }
    }

    #endregion
}