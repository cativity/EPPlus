﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB           EPPlus 5.2
 *************************************************************************************************/

using OfficeOpenXml.Drawing.Style.Effect;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx;

/// <summary>
/// A pareto line for a histogram chart
/// </summary>
public class ExcelChartExParetoLine : ExcelDrawingBorder
{
    private readonly ExcelChart _chart;

    internal ExcelChartExParetoLine(ExcelChart chart, XmlNamespaceManager nsm, XmlNode node)
        : base(chart, nsm, node, "cx:spPr/a:ln", new string[] { "spPr", "axisId" }) =>
        this._chart = chart;

    ExcelDrawingEffectStyle _effect;

    /// <summary>
    /// Effects
    /// </summary>
    public ExcelDrawingEffectStyle Effect => this._effect ??= new ExcelDrawingEffectStyle(this._chart, this.NameSpaceManager, this.TopNode, "cx:spPr/a:effectLst", this.SchemaNodeOrder);
}