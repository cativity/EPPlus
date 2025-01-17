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
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// Represents an individual datapoint in a chart
/// </summary>
public class ExcelChartDataPoint : XmlHelper, IDisposable, IDrawingStyleBase
{
    internal const string topNodePath = "c:dPt";
    ExcelChart _chart;

    internal ExcelChartDataPoint(ExcelChart chart, XmlNamespaceManager ns, XmlNode topNode)
        : base(ns, topNode)
    {
        this.Init(chart);
        this.Index = this.GetXmlNodeInt(indexPath);
    }

    internal ExcelChartDataPoint(ExcelChart chart, XmlNamespaceManager ns, XmlNode topNode, int index)
        : base(ns, topNode)
    {
        this.Init(chart);
        this.SetXmlNodeString(indexPath, index.ToString(CultureInfo.InvariantCulture));
        this.Bubble3D = false;
        this.Index = index;
    }

    private void Init(ExcelChart chart)
    {
        this._chart = chart;

        this.AddSchemaNodeOrder(new string[] { "idx", "invertIfNegative", "marker", "bubble3D", "explosion", "spPr", "pictureOptions", "extLst" },
                                ExcelDrawing._schemaNodeOrderSpPr);
    }

    const string indexPath = "c:idx/@val";

    /// <summary>
    /// The index of the datapoint
    /// </summary>
    public int Index { get; private set; }

    /// <summary>
    /// The sizes of the bubbles on the bubble chart
    /// </summary>
    public bool Bubble3D
    {
        get => this.GetXmlNodeBool("c:bubble3D/@val");
        set => this.SetXmlNodeString("c:bubble3D/@val", value.GetStringValueForXml());
    }

    /// <summary>
    /// Invert if negative. Default true.
    /// </summary>
    public bool InvertIfNegative
    {
        get => this.GetXmlNodeBool("c:invertIfNegative");
        set => this.SetXmlNodeString("c:invertIfNegative", value.GetStringValueForXml());
    }

    ExcelChartMarker _chartMarker;

    /// <summary>
    /// A reference to marker properties
    /// </summary>
    public ExcelChartMarker Marker => this._chartMarker ??= new ExcelChartMarker(this._chart, this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder);

    ExcelDrawingFill _fill;

    /// <summary>
    /// A reference to fill properties
    /// </summary>
    public ExcelDrawingFill Fill => this._fill ??= new ExcelDrawingFill(this._chart, this.NameSpaceManager, this.TopNode, "c:spPr", this.SchemaNodeOrder);

    ExcelDrawingBorder _line;

    /// <summary>
    /// A reference to line properties
    /// </summary>
    public ExcelDrawingBorder Border => this._line ??= new ExcelDrawingBorder(this._chart, this.NameSpaceManager, this.TopNode, "c:spPr/a:ln", this.SchemaNodeOrder);

    private ExcelDrawingEffectStyle _effect;

    /// <summary>
    /// A reference to line properties
    /// </summary>
    public ExcelDrawingEffectStyle Effect => this._effect ??= new ExcelDrawingEffectStyle(this._chart, this.NameSpaceManager, this.TopNode, "c:spPr/a:effectLst", this.SchemaNodeOrder);

    ExcelDrawing3D _threeD;

    /// <summary>
    /// 3D properties
    /// </summary>
    public ExcelDrawing3D ThreeD => this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, "c:spPr", this.SchemaNodeOrder);

    void IDrawingStyleBase.CreatespPr() => this.CreatespPrNode();

    /// <summary>
    /// Returns true if the datapoint has a marker
    /// </summary>
    /// <returns></returns>
    public bool HasMarker() => this.ExistsNode("c:marker");

    /// <summary>
    /// Dispose the object
    /// </summary>
    public void Dispose()
    {
        if (this._chart != null)
        {
            this._chart.Dispose();
        }

        this._chart = null;
        this._line = null;

        if (this._fill != null)
        {
            this._fill.Dispose();
        }

        this._fill = null;
    }
}