/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/08/2021         EPPlus Software AB       Table Styling - EPPlus 5.6
 *************************************************************************************************/

using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Style;

/// <summary>
/// A style element for a custom slicer style with band
/// </summary>
public class ExcelSlicerTableStyleElement : XmlHelper
{
    ExcelStyles _styles;

    internal ExcelSlicerTableStyleElement(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles, eTableStyleElement type)
        : base(nameSpaceManager, topNode)
    {
        this._styles = styles;
        this.Type = type;
    }

    ExcelDxfSlicerStyle _style;

    /// <summary>
    /// Access to style settings
    /// </summary>
    public ExcelDxfSlicerStyle Style
    {
        get { return this._style ??= this._styles.GetDxfSlicer(this.GetXmlNodeIntNull("@dxfId"), null); }
        internal set { this._style = value; }
    }

    /// <summary>
    /// The type of custom style element for a table style
    /// </summary>
    public eTableStyleElement Type { get; }

    internal virtual void CreateNode()
    {
        if (this.TopNode.LocalName != "tableStyleElement")
        {
            this.TopNode = this.CreateNode("d:tableStyleElement", false, true);
        }

        this.SetXmlNodeString("@type", this.Type.ToEnumString());
        this.SetXmlNodeInt("@dxfId", this.Style.DxfId);
    }
}