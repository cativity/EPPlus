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
using OfficeOpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.Style;

/// <summary>
/// Represents a style for a chart
/// </summary>
public class ExcelChartStyle : XmlHelper, IPictureRelationDocument
{
    ExcelChartStyleManager _manager;
    Dictionary<string, HashInfo> _hashes = new Dictionary<string, HashInfo>();

    internal ExcelChartStyle(XmlNamespaceManager nsm, XmlNode topNode, ExcelChartStyleManager manager)
        : base(nsm, topNode) =>
        this._manager = manager;

    ExcelChartStyleEntry _axisTitle;

    /// <summary>
    /// Default formatting for an axis title.
    /// </summary>
    public ExcelChartStyleEntry AxisTitle => this._axisTitle ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:axisTitle", this);

    ExcelChartStyleEntry _categoryAxis;

    /// <summary>
    /// Default formatting for a category axis
    /// </summary>
    public ExcelChartStyleEntry CategoryAxis => this._categoryAxis ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:categoryAxis", this);

    ExcelChartStyleEntry _chartArea;

    /// <summary>
    /// Default formatting for a chart area
    /// </summary>
    public ExcelChartStyleEntry ChartArea => this._chartArea ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:chartArea", this);

    ExcelChartStyleEntry _dataLabel;

    /// <summary>
    /// Default formatting for a data label
    /// </summary>
    public ExcelChartStyleEntry DataLabel => this._dataLabel ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataLabel", this);

    ExcelChartStyleEntry _dataLabelCallout;

    /// <summary>
    /// Default formatting for a data label callout
    /// </summary>
    public ExcelChartStyleEntry DataLabelCallout => this._dataLabelCallout ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataLabelCallout", this);

    ExcelChartStyleEntry _dataPoint;

    /// <summary>
    /// Default formatting for a data point on a 2-D chart of type column, bar, filled radar, stock, bubble, pie, doughnut, area and 3-D bubble.
    /// </summary>
    public ExcelChartStyleEntry DataPoint => this._dataPoint ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataPoint", this);

    ExcelChartStyleEntry _dataPoint3D;

    /// <summary>
    /// Default formatting for a data point on a 3-D chart of type column, bar, line, pie, area and surface.
    /// </summary>
    public ExcelChartStyleEntry DataPoint3D => this._dataPoint3D ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataPoint3D", this);

    ExcelChartStyleEntry _dataPointLine;

    /// <summary>
    /// Default formatting for a data point on a 2-D chart of type line, scatter and radar
    /// </summary>
    public ExcelChartStyleEntry DataPointLine => this._dataPointLine ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataPointLine", this);

    ExcelChartStyleEntry _dataPointMarker;

    /// <summary>
    /// Default formatting for a datapoint marker
    /// </summary>
    public ExcelChartStyleEntry DataPointMarker => this._dataPointMarker ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataPointMarker", this);

    ExcelChartStyleMarkerLayout _dataPointMarkerLayout;

    /// <summary>
    /// Extended marker properties for a datapoint 
    /// </summary>
    public ExcelChartStyleMarkerLayout DataPointMarkerLayout
    {
        get
        {
            if (this._dataPointMarkerLayout == null)
            {
                XmlNode? node = this.GetNode("cs:dataPointMarkerLayout");

                if (node == null)
                {
                    throw new InvalidOperationException("Invalid Chartstyle xml: dataPointMarkerLayout element missing");
                }

                this._dataPointMarkerLayout = new ExcelChartStyleMarkerLayout(this.NameSpaceManager, node);
            }

            return this._dataPointMarkerLayout;
        }
    }

    ExcelChartStyleEntry _dataPointWireframe;

    /// <summary>
    /// Default formatting for a datapoint on a surface wireframe chart
    /// </summary>
    public ExcelChartStyleEntry DataPointWireframe => this._dataPointWireframe ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataPointWireframe", this);

    ExcelChartStyleEntry _dataTable;

    /// <summary>
    /// Default formatting for a Data table
    /// </summary>
    public ExcelChartStyleEntry DataTable => this._dataTable ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataTable", this);

    ExcelChartStyleEntry _downBar;

    /// <summary>
    /// Default formatting for a downbar
    /// </summary>
    public ExcelChartStyleEntry DownBar => this._downBar ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:downBar", this);

    ExcelChartStyleEntry _dropLine;

    /// <summary>
    /// Default formatting for a dropline
    /// </summary>
    public ExcelChartStyleEntry DropLine => this._dropLine ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dropLine", this);

    ExcelChartStyleEntry _errorBar;

    /// <summary>
    /// Default formatting for an errorbar
    /// </summary>
    public ExcelChartStyleEntry ErrorBar => this._errorBar ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:errorBar", this);

    ExcelChartStyleEntry _floor;

    /// <summary>
    /// Default formatting for a floor
    /// </summary>
    public ExcelChartStyleEntry Floor => this._floor ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:floor", this);

    ExcelChartStyleEntry _gridlineMajor;

    /// <summary>
    /// Default formatting for a major gridline
    /// </summary>
    public ExcelChartStyleEntry GridlineMajor => this._gridlineMajor ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:gridlineMajor", this);

    ExcelChartStyleEntry _gridlineMinor;

    /// <summary>
    /// Default formatting for a minor gridline
    /// </summary>
    public ExcelChartStyleEntry GridlineMinor => this._gridlineMinor ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:gridlineMinor", this);

    ExcelChartStyleEntry _hiLoLine;

    /// <summary>
    /// Default formatting for a high low line
    /// </summary>
    public ExcelChartStyleEntry HighLowLine => this._hiLoLine ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:hiLoLine", this);

    ExcelChartStyleEntry _leaderLine;

    /// <summary>
    /// Default formatting for a leader line
    /// </summary>
    public ExcelChartStyleEntry LeaderLine => this._leaderLine ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:leaderLine", this);

    /// <summary>
    /// Default formatting for a legend
    /// </summary>
    ExcelChartStyleEntry _legend;

    /// <summary>
    /// Default formatting for a chart legend
    /// </summary>
    public ExcelChartStyleEntry Legend => this._legend ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:legend", this);

    ExcelChartStyleEntry _plotArea;

    /// <summary>
    /// Default formatting for a plot area in a 2D chart
    /// </summary>
    public ExcelChartStyleEntry PlotArea => this._plotArea ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:plotArea", this);

    ExcelChartStyleEntry _plotArea3D;

    /// <summary>
    /// Default formatting for a plot area in a 3D chart
    /// </summary>
    public ExcelChartStyleEntry PlotArea3D => this._plotArea3D ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:plotArea3D", this);

    ExcelChartStyleEntry _seriesAxis;

    /// <summary>
    /// Default formatting for a series axis
    /// </summary>
    public ExcelChartStyleEntry SeriesAxis => this._seriesAxis ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:seriesAxis", this);

    ExcelChartStyleEntry _seriesLine;

    /// <summary>
    /// Default formatting for a series line
    /// </summary>
    public ExcelChartStyleEntry SeriesLine => this._seriesLine ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:seriesLine", this);

    ExcelChartStyleEntry _title;

    /// <summary>
    /// Default formatting for a chart title
    /// </summary>
    public ExcelChartStyleEntry Title => this._title ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:title", this);

    ExcelChartStyleEntry _trendline;

    /// <summary>
    /// Default formatting for a trend line
    /// </summary>
    public ExcelChartStyleEntry Trendline => this._trendline ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:trendline", this);

    ExcelChartStyleEntry _trendlineLabel;

    /// <summary>
    /// Default formatting for a trend line label
    /// </summary>
    public ExcelChartStyleEntry TrendlineLabel => this._trendlineLabel ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:trendlineLabel", this);

    ExcelChartStyleEntry _upBar;

    /// <summary>
    /// Default formatting for a up bar
    /// </summary>
    public ExcelChartStyleEntry UpBar => this._upBar ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:upBar", this);

    ExcelChartStyleEntry _valueAxis;

    /// <summary>
    /// Default formatting for a value axis
    /// </summary>
    public ExcelChartStyleEntry ValueAxis => this._valueAxis ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:valueAxis", this);

    ExcelChartStyleEntry _wall;

    /// <summary>
    /// Default formatting for a wall
    /// </summary>
    public ExcelChartStyleEntry Wall => this._wall ??= new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:wall", this);

    /// <summary>
    /// The id of the chart style
    /// </summary>
    public int Id
    {
        get => this.GetXmlNodeInt("@id");
        internal set => this.SetXmlNodeString("@id", value.ToString(CultureInfo.InvariantCulture));
    }

    ExcelPackage IPictureRelationDocument.Package => this._manager._chart._drawings._package;

    Dictionary<string, HashInfo> IPictureRelationDocument.Hashes => this._hashes;

    ZipPackagePart IPictureRelationDocument.RelatedPart => this._manager.StylePart;

    Uri IPictureRelationDocument.RelatedUri => this._manager.StyleUri;
}