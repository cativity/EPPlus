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
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;
namespace OfficeOpenXml.Style.XmlAccess;

/// <summary>
/// Xml access class for border items
/// </summary>
public sealed class ExcelBorderItemXml : StyleXmlHelper
{
    internal ExcelBorderItemXml(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
    {
        this._borderStyle=ExcelBorderStyle.None;
        this._color = new ExcelColorXml(this.NameSpaceManager);
    }
    internal ExcelBorderItemXml(XmlNamespaceManager nsm, XmlNode topNode) :
        base(nsm, topNode)
    {
        if (topNode != null)
        {
            this._borderStyle = GetBorderStyle(this.GetXmlNodeString("@style"));
            this._color = new ExcelColorXml(nsm, topNode.SelectSingleNode(_colorPath, nsm));
            this.Exists = true;
        }
        else
        {
            this.Exists = false;
        }
    }

    private static ExcelBorderStyle GetBorderStyle(string style)
    {
        if(style=="")
        {
            return ExcelBorderStyle.None;
        }

        string sInStyle = style.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + style.Substring(1, style.Length - 1);
        try
        {
            return (ExcelBorderStyle)Enum.Parse(typeof(ExcelBorderStyle), sInStyle);
        }
        catch
        {
            return ExcelBorderStyle.None;
        }

    }
    ExcelBorderStyle _borderStyle = ExcelBorderStyle.None;
    /// <summary>
    /// Cell Border style
    /// </summary>
    public ExcelBorderStyle Style
    {
        get
        {
            return this._borderStyle;
        }
        set
        {
            this._borderStyle = value;
            this.Exists = true;
        }
    }
    ExcelColorXml _color = null;
    const string _colorPath = "d:color";
    /// <summary>
    /// The color of the line
    /// </summary>s
    public ExcelColorXml Color
    {
        get
        {
            return this._color;
        }
        internal set
        {
            this._color = value;
        }
    }
    internal override string Id
    {
        get 
        {
            if (this.Exists)
            {
                return this.Style + this.Color.Id;
            }
            else
            {
                return "None";
            }
        }
    }

    internal ExcelBorderItemXml Copy()
    {
        ExcelBorderItemXml? borderItem = new ExcelBorderItemXml(this.NameSpaceManager);
        borderItem.Style = this._borderStyle;
        borderItem.Color = this._color==null ? new ExcelColorXml(this.NameSpaceManager) { Auto = true } : this._color.Copy();
        return borderItem;
    }

    internal override XmlNode CreateXmlNode(XmlNode topNode)
    {
        this.TopNode = topNode;

        if (this.Style != ExcelBorderStyle.None)
        {
            this.SetXmlNodeString("@style", SetBorderString(this.Style));
            if (this.Color.Exists)
            {
                this.CreateNode(_colorPath);
                topNode.AppendChild(this.Color.CreateXmlNode(this.TopNode.SelectSingleNode(_colorPath, this.NameSpaceManager)));
            }
        }
        return this.TopNode;
    }

    private static string SetBorderString(ExcelBorderStyle Style)
    {
        string newName=Enum.GetName(typeof(ExcelBorderStyle), Style);
        return newName.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + newName.Substring(1, newName.Length - 1);
    }
    /// <summary>
    /// True if the record exists in the underlaying xml
    /// </summary>
    public bool Exists { get; private set; }
}