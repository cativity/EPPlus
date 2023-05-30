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

using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// Represents a up-down bar, dropline or hi-low line in a chart
/// </summary>
public class ExcelChartStyleItem : XmlHelper, IDrawingStyleBase
{
    ExcelChart _chart;
    //string _path;
    Action _removeMe;

    internal ExcelChartStyleItem(XmlNamespaceManager nsm, XmlNode topNode, ExcelChart chart, string path, Action removeMe)
        : base(nsm, topNode)
    {
        this._chart = chart;
        //this._path = path;
        this.AddSchemaNodeOrder(chart._chartXmlHelper.SchemaNodeOrder, ExcelDrawing._schemaNodeOrderSpPr);
        this.TopNode = this.CreateNode(path);
        this._removeMe = removeMe;
    }

    ExcelDrawingFill _fill;

    /// <summary>
    /// Access to fill properties
    /// </summary>
    public ExcelDrawingFill Fill
    {
        get { return this._fill ??= new ExcelDrawingFill(this._chart, this.NameSpaceManager, this.TopNode, "c:spPr", this.SchemaNodeOrder); }
    }

    ExcelDrawingBorder _border;

    /// <summary>
    /// Access to border properties
    /// </summary>
    public ExcelDrawingBorder Border
    {
        get { return this._border ??= new ExcelDrawingBorder(this._chart, this.NameSpaceManager, this.TopNode, "c:spPr/a:ln", this.SchemaNodeOrder); }
    }

    ExcelDrawingEffectStyle _effect;

    /// <summary>
    /// Effects
    /// </summary>
    public ExcelDrawingEffectStyle Effect
    {
        get
        {
            return this._effect ??= new ExcelDrawingEffectStyle(this._chart, this.NameSpaceManager, this.TopNode, "c:spPr/a:effectLst", this.SchemaNodeOrder);
        }
    }

    ExcelDrawing3D _threeD;

    /// <summary>
    /// 3D properties
    /// </summary>
    public ExcelDrawing3D ThreeD
    {
        get { return this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, "c:spPr", this.SchemaNodeOrder); }
    }

    void IDrawingStyleBase.CreatespPr()
    {
        this.CreatespPrNode();
    }

    /// <summary>
    /// Removes the item
    /// </summary>
    public void Remove()
    {
        _ = this.TopNode.ParentNode.RemoveChild(this.TopNode);
        this._removeMe();
    }
}