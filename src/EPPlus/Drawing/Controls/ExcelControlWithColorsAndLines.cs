﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    11/24/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/

using System.Xml;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.Drawing.Controls;

/// <summary>
/// An abstract class used by form controls with color and line settings
/// </summary>
public abstract class ExcelControlWithColorsAndLines : ExcelControlWithText
{
    internal ExcelControlWithColorsAndLines(ExcelDrawings drawings, XmlElement drawNode, string name, ExcelGroupShape parent)
        : base(drawings, drawNode, name, parent) =>
        this.SetSize(90, 30); //Default size        

    internal ExcelControlWithColorsAndLines(ExcelDrawings drawings,
                                            XmlNode drawingNode,
                                            ControlInternal control,
                                            ZipPackagePart part,
                                            XmlDocument ctrlPropXml,
                                            ExcelGroupShape parent = null)
        : base(drawings, drawingNode, control, part, ctrlPropXml, parent)
    {
    }

    /// <summary>
    /// Fill settings for the control
    /// </summary>
    public ExcelVmlDrawingFill Fill => this._vml.GetFill();

    ExcelVmlDrawingBorder _border;

    /// <summary>
    /// Border settings for the control
    /// </summary>
    public ExcelVmlDrawingBorder Border => this._border ??= new ExcelVmlDrawingBorder(this._drawings, this._vml.NameSpaceManager, this._vml.TopNode, this._vml.SchemaNodeOrder);

    internal override void UpdateXml()
    {
        base.UpdateXml();
        this.Border.UpdateXml();
    }
}