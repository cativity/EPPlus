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
  02/26/2021         EPPlus Software AB       Modified to work with dxf styling for tables
 *************************************************************************************************/

using System;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf;

/// <summary>
/// Differential formatting record used in conditional formatting
/// </summary>
public class ExcelDxfSlicerStyle : ExcelDxfStyleFont
{
    internal ExcelDxfSlicerStyle(XmlNamespaceManager nameSpaceManager,
                                 XmlNode topNode,
                                 ExcelStyles styles,
                                 Action<eStyleClass, eStyleProperty, object> callback)
        : base(nameSpaceManager, topNode, styles, callback)
    {
    }

    internal override DxfStyleBase Clone()
    {
        ExcelDxfSlicerStyle? s = new ExcelDxfSlicerStyle(this._helper.NameSpaceManager, null, this._styles, this._callback)
        {
            Font = (ExcelDxfFont)this.Font.Clone(), Fill = (ExcelDxfFill)this.Fill.Clone(), Border = (ExcelDxfBorderBase)this.Border.Clone(),
        };

        return s;
    }
}