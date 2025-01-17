﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/21/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/

using OfficeOpenXml.Packaging;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls;

/// <summary>
/// Represents a label form control.
/// </summary>
public class ExcelControlLabel : ExcelControlWithText
{
    internal ExcelControlLabel(ExcelDrawings drawings, XmlElement drawNode, string name, ExcelGroupShape parent = null)
        : base(drawings, drawNode, name, parent) =>
        this.SetSize(150, 30); //Default size

    internal ExcelControlLabel(ExcelDrawings drawings,
                               XmlNode drawNode,
                               ControlInternal control,
                               ZipPackagePart part,
                               XmlDocument controlPropertiesXml,
                               ExcelGroupShape parent = null)
        : base(drawings, drawNode, control, part, controlPropertiesXml, parent)
    {
    }

    /// <summary>
    /// The type of form control
    /// </summary>
    public override eControlType ControlType => eControlType.Label;
}