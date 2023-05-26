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
using System.Linq;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// Base class for standard charts with line properties.
/// </summary>
public abstract class ExcelStandardChartWithLines : ExcelChartStandard, IDrawingDataLabel
{
    internal ExcelStandardChartWithLines(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
        base(drawings, node, uriChart, part, chartXml, chartNode, parent)
    {
    }

    internal ExcelStandardChartWithLines(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent = null) :
        base(topChart, chartNode, parent)
    {
    }
    internal ExcelStandardChartWithLines(ExcelDrawings drawings, XmlNode node, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml, ExcelGroupShape parent = null) :
        base(drawings, node, type, topChart, PivotTableSource, chartXml, parent)
    {

    }
    string MARKER_PATH = "c:marker/@val";
    /// <summary>
    /// If the series has markers
    /// </summary>
    public bool Marker
    {
        get
        {
            return this._chartXmlHelper.GetXmlNodeBool(this.MARKER_PATH, false);
        }
        set
        {
            this._chartXmlHelper.SetXmlNodeBool(this.MARKER_PATH, value, false);
        }
    }

    string SMOOTH_PATH = "c:smooth/@val";
    /// <summary>
    /// If the series has smooth lines
    /// </summary>
    public bool Smooth
    {
        get
        {
            return this._chartXmlHelper.GetXmlNodeBool(this.SMOOTH_PATH, false);
        }
        set
        {
            if (this.ChartType == eChartType.Line3D)
            {
                throw new ArgumentException("Smooth", "Smooth does not apply to 3d line charts");
            }

            this._chartXmlHelper.SetXmlNodeBool(this.SMOOTH_PATH, value);
        }
    }
    //string _chartTopPath = "c:chartSpace/c:chart/c:plotArea/{0}";
    ExcelChartDataLabel _dataLabel = null;
    /// <summary>
    /// Access to datalabel properties
    /// </summary>
    public ExcelChartDataLabel DataLabel
    {
        get
        {
            return this._dataLabel ??= new ExcelChartDataLabelStandard(this, this.NameSpaceManager, this.ChartNode, "dLbls", this._chartXmlHelper.SchemaNodeOrder);
        }
    }
    /// <summary>
    /// If the chart has datalabel
    /// </summary>
    public bool HasDataLabel
    {
        get
        {
            return this.ChartNode.SelectSingleNode("c:dLbls", this.NameSpaceManager) != null;
        }
    }
    const string _gapWidthPath = "c:upDownBars/c:gapWidth/@val";
    /// <summary>
    /// The gap width between the up and down bars
    /// </summary>
    public double? UpDownBarGapWidth
    {
        get
        {
            return this._chartXmlHelper.GetXmlNodeIntNull(_gapWidthPath);
        }
        set
        {
            if (value == null)
            {
                this._chartXmlHelper.DeleteNode(_gapWidthPath, true);
            }
            if (value < 0 || value > 500)
            {
                throw (new ArgumentOutOfRangeException("GapWidth ranges between 0 and 500"));
            }

            this._chartXmlHelper.SetXmlNodeString(_gapWidthPath, value.Value.ToString(CultureInfo.InvariantCulture));
        }
    }
    ExcelChartStyleItem _upBar = null;
    const string _upBarPath = "c:upDownBars/c:upBars";
    /// <summary>
    /// Format the up bars on the chart
    /// </summary>
    public ExcelChartStyleItem UpBar
    {
        get
        {
            return this._upBar;
        }
    }
    ExcelChartStyleItem _downBar = null;
    const string _downBarPath = "c:upDownBars/c:downBars";
    /// <summary>
    /// Format the down bars on the chart
    /// </summary>
    public ExcelChartStyleItem DownBar
    {
        get
        {
            return this._downBar;
        }
    }
    ExcelChartStyleItem _hiLowLines = null;
    const string _hiLowLinesPath = "c:hiLowLines";
    /// <summary>
    /// Format the high-low lines for the series.
    /// </summary>
    public ExcelChartStyleItem HighLowLine
    {
        get
        {
            return this._hiLowLines;
        }
    }
    ExcelChartStyleItem _dropLines = null;
    const string _dropLinesPath = "c:dropLines";

    /// <summary>
    /// Format the drop lines for the series.
    /// </summary>
    public ExcelChartStyleItem DropLine
    {
        get
        {
            return this._dropLines;
        }
    }
    /// <summary>
    /// Adds up and/or down bars to the chart.        
    /// </summary>
    /// <param name="upBars">Adds up bars if up bars does not exist.</param>
    /// <param name="downBars">Adds down bars if down bars does not exist.</param>
    public void AddUpDownBars(bool upBars = true, bool downBars = true)
    {
        if (upBars && this._upBar == null)
        {
            this._upBar = new ExcelChartStyleItem(this.NameSpaceManager, this.ChartNode, this, _upBarPath, this.RemoveUpBar);
            ExcelChart? chart = this._topChart ?? this;
            chart.ApplyStyleOnPart(this._upBar, chart.StyleManager?.Style?.UpBar);
        }
        if (downBars && this._downBar == null)
        {
            this._downBar = new ExcelChartStyleItem(this.NameSpaceManager, this.ChartNode, this, _downBarPath, this.RemoveDownBar);
            ExcelChart? chart = this._topChart ?? this;
            chart.ApplyStyleOnPart(this._upBar, chart.StyleManager?.Style?.DownBar);
        }
    }
    /// <summary>
    /// Adds droplines to the chart.        
    /// </summary>
    public ExcelChartStyleItem AddDropLines()
    {
        if (this._dropLines == null)
        {
            this._dropLines = new ExcelChartStyleItem(this.NameSpaceManager, this.ChartNode, this, _dropLinesPath, this.RemoveDropLines);
            ExcelChart? chart = this._topChart ?? this;
            chart.ApplyStyleOnPart(this._upBar, chart.StyleManager?.Style?.DropLine);
        }
        return this._dropLines;
    }
    /// <summary>
    /// Adds High-Low lines to the chart.        
    /// </summary>
    public ExcelChartStyleItem AddHighLowLines()
    {
        if (this._hiLowLines == null)
        {
            this._hiLowLines = new ExcelChartStyleItem(this.NameSpaceManager, this.ChartNode, this, _hiLowLinesPath, this.RemoveHiLowLines);
            ExcelChart? chart = this._topChart ?? this;
            chart.ApplyStyleOnPart(this._upBar, chart.StyleManager?.Style?.HighLowLine);
        }
        return this.HighLowLine;
    }
    //TODO: Consider adding this method later (for all charts with datalabels)
    ///// <summary>
    ///// Adds datalabels to the chart
    ///// </summary>
    ///// <param name="position">The position of the datalabels</param>
    ///// <returns></returns>
    //public ExcelChartDataLabel AddDataLabels(eLabelPosition position=eLabelPosition.Center)
    //{
    //    DataLabel.Position = position;
    //    var chart = _topChart ?? this;
    //    foreach (var serie in chart.Series)
    //    {
    //        if (serie is IDrawingSerieDataLabel dl)
    //            dl.DataLabel.Position = position;
    //        if (chart.StyleManager.StylePart != null)
    //        {
    //            chart.StyleManager.ApplyStyle(serie, chart.StyleManager.Style.DataLabel);
    //        }

    //    }
    //    return DataLabel;
    //}
    internal override eChartType GetChartType(string name)
    {
        if (name == "lineChart")
        {
            if (this.Marker)
            {
                if (this.Grouping == eGrouping.Stacked)
                {
                    return eChartType.LineMarkersStacked;
                }
                else if (this.Grouping == eGrouping.PercentStacked)
                {
                    return eChartType.LineMarkersStacked100;
                }
                else
                {
                    return eChartType.LineMarkers;
                }
            }
            else
            {
                if (this.Grouping == eGrouping.Stacked)
                {
                    return eChartType.LineStacked;
                }
                else if (this.Grouping == eGrouping.PercentStacked)
                {
                    return eChartType.LineStacked100;
                }
                else
                {
                    return eChartType.Line;
                }
            }
        }
        else if (name == "line3DChart")
        {
            return eChartType.Line3D;
        }
        return base.GetChartType(name);
    }
    internal override void InitSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list = null)
    {
        base.InitSeries(chart, ns, node, isPivot, list);
        this.AddSchemaNodeOrder(this.SchemaNodeOrder, new string[] { "gapWidth", "upbars", "downbars" });
        this.Series.Init(chart, ns, node, isPivot, base.Series._list);

        //Up bars
        if (this._upBar==null && this.ExistsNode(node, _upBarPath))
        {
            this._upBar = new ExcelChartStyleItem(ns, node, this, _upBarPath, this.RemoveUpBar);
        }

        //Down bars
        if (this._downBar == null && this.ExistsNode(node, _downBarPath))
        {
            this._downBar = new ExcelChartStyleItem(ns, node, this, _downBarPath, this.RemoveDownBar);
        }

        //Drop lines
        if (this._dropLines == null && this.ExistsNode(node, _dropLinesPath))
        {
            this._dropLines = new ExcelChartStyleItem(ns, node, this, _dropLinesPath, this.RemoveDropLines);
        }

        //High / low lines
        if (this._hiLowLines == null && this.ExistsNode(node, _hiLowLinesPath))
        {
            this._hiLowLines = new ExcelChartStyleItem(ns, node, this, _hiLowLinesPath, this.RemoveHiLowLines);
        }


    }

    /// <summary>
    /// The series for the chart
    /// </summary>
    public new ExcelChartSeries<ExcelLineChartSerie> Series
    {
        get;
    } = new ExcelChartSeries<ExcelLineChartSerie>();
    #region Remove Line/Bar
    private void RemoveUpBar()
    {
        this._upBar = null;
    }
    private void RemoveDownBar()
    {
        this._downBar = null;
    }
    private void RemoveDropLines()
    {
        this._dropLines = null;
    }
    private void RemoveHiLowLines()
    {
        this._hiLowLines = null;
    }
    #endregion

}

/// <summary>
/// Provides access to line chart specific properties
/// </summary>
public class ExcelLineChart : ExcelStandardChartWithLines
{
    #region "Constructors"
    internal ExcelLineChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
        base(drawings, node, uriChart, part, chartXml, chartNode, parent)
    {
    }

    internal ExcelLineChart (ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent = null) :
        base(topChart, chartNode, parent)
    {
    }
    internal ExcelLineChart(ExcelDrawings drawings, XmlNode node, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml, ExcelGroupShape parent = null) :
        base(drawings, node, type, topChart, PivotTableSource, chartXml, parent)
    {
        if (type != eChartType.Line3D)
        {
            this.Smooth = false;
        }
    }
    #endregion
}