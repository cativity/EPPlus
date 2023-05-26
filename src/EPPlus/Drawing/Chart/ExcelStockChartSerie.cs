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
using System.Text;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// A serie for a scatter chart
/// </summary>
public sealed class ExcelStockChartSerie : ExcelChartSerieWithErrorBars, IDrawingSerieDataLabel, IDrawingChartMarker
{
    /// <summary>
    /// Default constructor
    /// </summary>
    /// <param name="chart">The chart</param>
    /// <param name="ns">Namespacemanager</param>
    /// <param name="node">Topnode</param>
    /// <param name="isPivot">Is pivotchart</param>
    internal ExcelStockChartSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot) :
        base(chart, ns, node, isPivot)
    {
        this.Marker.Style = eMarkerStyle.None;
        this.Smooth = 0;
        this.Border.LineCap = eLineCap.Round;
        this.Border.Fill.Style = eFillStyle.NoFill;            
    }

    ExcelChartSerieDataLabel _dataLabel = null;
    /// <summary>
    /// Data label properties
    /// </summary>
    public ExcelChartSerieDataLabel DataLabel
    {
        get
        {
            return this._dataLabel ??= new ExcelChartSerieDataLabel(this._chart, this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder);
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
    const string smoothPath = "c:smooth/@val";
    /// <summary>
    /// Smooth for scattercharts
    /// </summary>
    public int Smooth
    {
        get
        {
            return this.GetXmlNodeInt(smoothPath);
        }
        internal set
        {
            this.SetXmlNodeString(smoothPath, value.ToString());
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
            if (this.IsMarkersAllowed() == false)
            {
                return null;
            }

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
            return this.Marker.Style != eMarkerStyle.None;
        }
        return false;
    }
    private bool IsMarkersAllowed()
    {
        eChartType type = this._chart.ChartType;
        //if (type == eChartType.XYScatterLinesNoMarkers || type == eChartType.XYScatterSmoothNoMarkers)
        //{
        //    return false;
        //}
        return true;
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
    /// <summary>
    /// Line color.
    /// </summary>
    ///
    /// <value>
    /// The color of the line.
    /// </value>
    [Obsolete("Please use Border.Fill.Color property")]
    public Color LineColor
    {
        get
        {
            if (this.Border.Fill.Style == eFillStyle.SolidFill && this.Border.Fill.SolidFill.Color.ColorType == eDrawingColorType.Rgb)
            {
                return this.Border.Fill.Color;
            }
            else
            {
                return Color.Black;
            }
        }
        set
        {
            this.Border.Fill.Color = value;
        }
    }
    /// <summary>
    /// Gets or sets the size of the marker.
    /// </summary>
    ///
    /// <remarks>
    /// value between 2 and 72.
    /// </remarks>
    ///
    /// <value>
    /// The size of the marker.
    /// </value>
    [Obsolete("Please use Marker.Size")]
    public int MarkerSize
    {
        get
        {

            int size = this.Marker.Size;
            if (size == 0)
            {
                return 5;
            }
            else
            {
                return size;
            }
        }
        set
        {
            this.Marker.Size = value;
        }
    }
    /// <summary>
    /// Marker color.
    /// </summary>
    /// <value>
    /// The color of the Marker.
    /// </value>
    [Obsolete("Please use Marker.Fill")]
    public Color MarkerColor
    {
        get
        {
            if (this.Marker.Fill.Style == eFillStyle.SolidFill && this.Marker.Fill.SolidFill.Color.ColorType == eDrawingColorType.Rgb)
            {
                return this.Marker.Fill.Color;
            }
            else
            {
                return Color.Black;
            }
        }
        set
        {
            this.Marker.Fill.Color=value;
        }
    }

    /// <summary>
    /// Gets or sets the width of the line in pt.
    /// </summary>
    ///
    /// <value>
    /// The width of the line.
    /// </value>
    [Obsolete("Please use Border.Width")]
    public double LineWidth
    {
        get
        {
            double width = this.Border.Width;
            if (width == 0)
            {
                return 2.25;
            }
            else
            {
                return width;
            }
        }
        set
        {
            this.Border.Width = value;
        }
    }
    /// <summary>
    /// Marker Line color.
    /// (not to be confused with LineColor)
    /// </summary>
    ///
    /// <value>
    /// The color of the Marker line.
    /// </value>
    [Obsolete("Please use Marker.Border.Fill.Color")]
    public Color MarkerLineColor
    {
        get
        {                
            if (this.Marker.Border.Fill.Style==eFillStyle.SolidFill && this.Marker.Border.Fill.SolidFill.Color.ColorType==eDrawingColorType.Rgb)
            {
                return this.Marker.Border.Fill.Color;
            }
            else
            {
                return Color.Black;
            }
        }
        set
        {
            this.Marker.Border.Fill.Color = value;
        }
    }       
}