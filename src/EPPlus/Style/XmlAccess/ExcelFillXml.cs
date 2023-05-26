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
namespace OfficeOpenXml.Style.XmlAccess;

/// <summary>
/// Xml access class for fills
/// </summary>
public class ExcelFillXml : StyleXmlHelper 
{
    internal ExcelFillXml(XmlNamespaceManager nameSpaceManager)
        : base(nameSpaceManager)
    {
        this._fillPatternType = ExcelFillStyle.None;
        this._backgroundColor = new ExcelColorXml(this.NameSpaceManager);
        this._patternColor = new ExcelColorXml(this.NameSpaceManager);
    }
    internal ExcelFillXml(XmlNamespaceManager nsm, XmlNode topNode):
        base(nsm, topNode)
    {
        this.PatternType = GetPatternType(this.GetXmlNodeString(fillPatternTypePath));
        this._backgroundColor = new ExcelColorXml(nsm, topNode.SelectSingleNode(_backgroundColorPath, nsm));
        this._patternColor = new ExcelColorXml(nsm, topNode.SelectSingleNode(_patternColorPath, nsm));
    }

    private static ExcelFillStyle GetPatternType(string patternType)
    {
        if (patternType == "")
        {
            return ExcelFillStyle.None;
        }

        patternType = patternType.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + patternType.Substring(1, patternType.Length - 1);
        try
        {
            return (ExcelFillStyle)Enum.Parse(typeof(ExcelFillStyle), patternType);
        }
        catch
        {
            return ExcelFillStyle.None;
        }
    }
    internal override string Id
    {
        get
        {
            return this.PatternType + this.PatternColor.Id + this.BackgroundColor.Id;
        }
    }
    #region Public Properties
    const string fillPatternTypePath = "d:patternFill/@patternType";
    internal ExcelFillStyle _fillPatternType;
    /// <summary>
    /// Cell fill pattern style
    /// </summary>
    public ExcelFillStyle PatternType
    {
        get
        {
            return this._fillPatternType;
        }
        set
        {
            this._fillPatternType=value;
        }
    }
    internal ExcelColorXml _patternColor = null;
    const string _patternColorPath = "d:patternFill/d:bgColor";
    /// <summary>
    /// Pattern color
    /// </summary>
    public ExcelColorXml PatternColor
    {
        get
        {
            return this._patternColor;
        }
        internal set
        {
            this._patternColor = value;
        }
    }
    internal ExcelColorXml _backgroundColor = null;
    const string _backgroundColorPath = "d:patternFill/d:fgColor";
    /// <summary>
    /// Cell background color 
    /// </summary>
    public ExcelColorXml BackgroundColor
    {
        get
        {
            return this._backgroundColor;
        }
        internal set
        {
            this._backgroundColor=value;
        }
    }
    #endregion


    //internal Fill Copy()
    //{
    //    Fill newFill = new Fill(NameSpaceManager, TopNode.Clone());
    //    return newFill;
    //}

    internal virtual ExcelFillXml Copy()
    {
        ExcelFillXml newFill = new ExcelFillXml(this.NameSpaceManager);
        newFill.PatternType = this._fillPatternType;
        newFill.BackgroundColor = this._backgroundColor.Copy();
        newFill.PatternColor = this._patternColor.Copy();
        return newFill;
    }

    internal override XmlNode CreateXmlNode(XmlNode topNode)
    {
        this.TopNode = topNode;
        this.SetXmlNodeString(fillPatternTypePath, SetPatternString(this._fillPatternType));
        if (this.PatternType != ExcelFillStyle.None)
        {
            XmlNode pattern = topNode.SelectSingleNode(fillPatternTypePath, this.NameSpaceManager);
            if (this.BackgroundColor.Exists)
            {
                this.CreateNode(_backgroundColorPath);
                this.BackgroundColor.CreateXmlNode(topNode.SelectSingleNode(_backgroundColorPath, this.NameSpaceManager));
                if (this.PatternColor.Exists)
                {
                    this.CreateNode(_patternColorPath);
                    this.PatternColor.CreateXmlNode(topNode.SelectSingleNode(_patternColorPath, this.NameSpaceManager));
                }
            }
        }
        return topNode;
    }

    private static string SetPatternString(ExcelFillStyle pattern)
    {
        string newName = Enum.GetName(typeof(ExcelFillStyle), pattern);
        return newName.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + newName.Substring(1, newName.Length - 1);
    }
}