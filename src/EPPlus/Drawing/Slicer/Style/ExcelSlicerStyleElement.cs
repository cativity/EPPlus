﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2021         EPPlus Software AB       Table Styling - EPPlus 5.6
 *************************************************************************************************/

using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer.Style;

/// <summary>
/// A style element for a custom slicer style 
/// </summary>
public class ExcelSlicerStyleElement : XmlHelper
{
    ExcelStyles _styles;

    internal ExcelSlicerStyleElement(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles, eSlicerStyleElement type)
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
        get => this._style ??= this._styles.GetDxfSlicer(this.GetXmlNodeIntNull("@dxfId"));
        internal set => this._style = value;
    }

    /// <summary>
    /// The type of the slicer element that this style is applied to.
    /// </summary>
    public eSlicerStyleElement Type { get; }

    internal virtual void CreateNode()
    {
        if (this.TopNode.LocalName != "slicerStyleElement")
        {
            this.TopNode = this.CreateNode("x14:slicerStyleElements");
            this.TopNode = this.CreateNode("x14:slicerStyleElement", false, true);
        }

        this.SetXmlNodeString("@type", this.Type.ToEnumString());
        this.SetXmlNodeInt("@dxfId", this.Style.DxfId);
    }
}