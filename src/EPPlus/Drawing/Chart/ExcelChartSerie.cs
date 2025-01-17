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
using System.Text;
using System.Xml;
using System.Linq;
using OfficeOpenXml.Core.CellStore;
using System.Globalization;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// Base class for chart series for standard charts
/// </summary>
public abstract class ExcelChartSerie : XmlHelper, IDrawingStyleBase
{
    internal ExcelChart _chart;
    string _prefix;

    internal ExcelChartSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, string prefix = "c")
        : base(ns, node)
    {
        this._chart = chart;
        this._prefix = prefix;
    }

    /// <summary>
    /// The header for the chart serie
    /// </summary>
    public abstract string Header { get; set; }

    /// <summary>
    /// Literals for the Y serie, if the literal values are numeric
    /// </summary>
    public double[] NumberLiteralsY { get; protected set; }

    /// <summary>
    /// Literals for the X serie, if the literal values are numeric
    /// </summary>
    public double[] NumberLiteralsX { get; protected set; }

    /// <summary>
    /// Literals for the X serie, if the literal values are strings
    /// </summary>
    public string[] StringLiteralsX { get; protected set; }

    void IDrawingStyleBase.CreatespPr() => this.CreatespPrNode();

    /// <summary>
    /// The header address for the serie.
    /// </summary>
    public abstract ExcelAddressBase HeaderAddress { get; set; }

    /// <summary>
    /// The address for the vertical series.
    /// </summary>
    public abstract string Series { get; set; }

    /// <summary>
    /// The address for the horizontal series.
    /// </summary>
    public abstract string XSeries { get; set; }

    ExcelDrawingFill _fill;

    /// <summary>
    /// Access to fill properties
    /// </summary>
    public ExcelDrawingFill Fill => this._fill ??= new ExcelDrawingFill(this._chart, this.NameSpaceManager, this.TopNode, $"{this._prefix}:spPr", this.SchemaNodeOrder);

    ExcelDrawingBorder _border;

    /// <summary>
    /// Access to border properties
    /// </summary>
    public ExcelDrawingBorder Border => this._border ??= new ExcelDrawingBorder(this._chart, this.NameSpaceManager, this.TopNode, $"{this._prefix}:spPr/a:ln", this.SchemaNodeOrder);

    ExcelDrawingEffectStyle _effect;

    /// <summary>
    /// Effects
    /// </summary>
    public ExcelDrawingEffectStyle Effect =>
        this._effect ??= new ExcelDrawingEffectStyle(this._chart,
                                                     this.NameSpaceManager,
                                                     this.TopNode,
                                                     $"{this._prefix}:spPr/a:effectLst",
                                                     this.SchemaNodeOrder);

    ExcelDrawing3D _threeD;

    /// <summary>
    /// 3D properties
    /// </summary>
    public ExcelDrawing3D ThreeD => this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, $"{this._prefix}:spPr", this.SchemaNodeOrder);

    /// <summary>
    /// Number of items in the serie.
    /// </summary>
    public abstract int NumberOfItems { get; }

    /// <summary>
    /// A collection of trend lines for the chart serie.
    /// </summary>
    public abstract ExcelChartTrendlineCollection TrendLines { get; }

    internal abstract void SetID(string id);

    internal string ToFullAddress(string value)
    {
        if (ExcelCellBase.IsValidAddress(value))
        {
            return ExcelCellBase.GetFullAddress(this._chart.WorkSheet.Name, value);
        }
        else
        {
            return value;
        }
    }
}