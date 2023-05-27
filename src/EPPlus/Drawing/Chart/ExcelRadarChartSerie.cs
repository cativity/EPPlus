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
using System;
using System.Globalization;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// A serie for a scatter chart
/// </summary>
public sealed class ExcelRadarChartSerie : ExcelChartStandardSerie, IDrawingSerieDataLabel, IDrawingChartMarker, IDrawingChartDataPoints
{
    /// <summary>
    /// Default constructor
    /// </summary>
    /// <param name="chart">The chart</param>
    /// <param name="ns">Namespacemanager</param>
    /// <param name="node">Topnode</param>
    /// <param name="isPivot">Is pivotchart</param>
    internal ExcelRadarChartSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot) :
        base(chart, ns, node, isPivot)
    {
        if (chart.ChartType == eChartType.RadarMarkers)
        {
            this.Marker.Style = eMarkerStyle.Square;
        }
        else if(chart.ChartType == eChartType.Radar)
        {
            this.Marker.Style = eMarkerStyle.None;
        }
    }
    ExcelChartSerieDataLabel _DataLabel = null;
    /// <summary>
    /// Datalabel
    /// </summary>
    public ExcelChartSerieDataLabel DataLabel
    {
        get
        {
            return this._DataLabel ??= new ExcelChartSerieDataLabel(this._chart, this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder);
        }
    }
    /// <summary>
    /// If the chart has datalabel
    /// </summary>
    public bool HasDataLabel
    {
        get
        {
            return this.TopNode.SelectSingleNode("c:dLbls", this.NameSpaceManager) != null;
        }
    }
    const string markerPath = "c:marker/c:symbol/@val";
    ExcelChartMarker _chartMarker = null;
    /// <summary>
    /// A reference to marker properties
    /// </summary>
    public ExcelChartMarker Marker
    {
        get
        {
            //if (IsMarkersAllowed() == false) return null;
            return this._chartMarker ??= new ExcelChartMarker(this._chart, this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder);
        }
    }
    /// <summary>
    /// If the serie has markers
    /// </summary>
    /// <returns>True if serie has markers</returns>
    public bool HasMarker()
    {
        if (this.IsMarkersAllowed())
        {
            return this.ExistsNode("c:marker");
        }
        return false;
    }
    private bool IsMarkersAllowed()
    {
        if (this._chart.ChartType == eChartType.RadarMarkers)
        {
            return true;
        }
        return false;
    }
    ExcelChartDataPointCollection _dataPoints = null;
    /// <summary>
    /// A collection of the individual datapoints
    /// </summary>
    public ExcelChartDataPointCollection DataPoints
    {
        get
        {
            return this._dataPoints ??= new ExcelChartDataPointCollection(this._chart, this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder);
        }
    }

    const string MARKERSIZE_PATH = "c:marker/c:size/@val";
    /// <summary>
    /// The size of a markers
    /// </summary>
    [Obsolete("Please use Marker.Size")]
    public int MarkerSize
    {
        get
        {
            return this.GetXmlNodeInt(MARKERSIZE_PATH);
        }
        set
        {
            if (value < 2 && value > 72)
            {
                throw new ArgumentOutOfRangeException("MarkerSize out of range. Range from 2-72 allowed.");
            }

            this.SetXmlNodeString(MARKERSIZE_PATH, value.ToString(CultureInfo.InvariantCulture));
        }
    }

}