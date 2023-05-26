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

namespace OfficeOpenXml.Drawing.Chart.Style
{
    /// <summary>
    /// Represents a style for a chart
    /// </summary>
    public class ExcelChartStyle : XmlHelper, IPictureRelationDocument
    {
        ExcelChartStyleManager _manager;
        Dictionary<string, HashInfo> _hashes = new Dictionary<string, HashInfo>();
        internal ExcelChartStyle(XmlNamespaceManager nsm, XmlNode topNode, ExcelChartStyleManager manager) : base(nsm, topNode)
        {
            this._manager = manager;            
        }
        ExcelChartStyleEntry _axisTitle = null;
        /// <summary>
        /// Default formatting for an axis title.
        /// </summary>
        public ExcelChartStyleEntry AxisTitle
        {
            get
            {
                if (this._axisTitle == null)
                {
                    this._axisTitle = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:axisTitle", this);
                }
                return this._axisTitle;
            }
        }
        ExcelChartStyleEntry _categoryAxis = null;
        /// <summary>
        /// Default formatting for a category axis
        /// </summary>
        public ExcelChartStyleEntry CategoryAxis
        {
            get
            {
                if (this._categoryAxis == null)
                {
                    this._categoryAxis = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:categoryAxis", this);
                }
                return this._categoryAxis;
            }
        }
        ExcelChartStyleEntry _chartArea = null;
        /// <summary>
        /// Default formatting for a chart area
        /// </summary>
        public ExcelChartStyleEntry ChartArea
        {
            get
            {
                if (this._chartArea == null)
                {
                    this._chartArea = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:chartArea", this);
                }
                return this._chartArea;
            }
        }
        ExcelChartStyleEntry _dataLabel = null;
        /// <summary>
        /// Default formatting for a data label
        /// </summary>
        public ExcelChartStyleEntry DataLabel
        {
            get
            {
                if (this._dataLabel == null)
                {
                    this._dataLabel = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataLabel", this);
                }
                return this._dataLabel;
            }
        }
        ExcelChartStyleEntry _dataLabelCallout = null;
        /// <summary>
        /// Default formatting for a data label callout
        /// </summary>
        public ExcelChartStyleEntry DataLabelCallout
        {
            get
            {
                if (this._dataLabelCallout == null)
                {
                    this._dataLabelCallout = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataLabelCallout", this);
                }
                return this._dataLabelCallout;
            }
        }
        ExcelChartStyleEntry _dataPoint = null;
        /// <summary>
        /// Default formatting for a data point on a 2-D chart of type column, bar,	filled radar, stock, bubble, pie, doughnut, area and 3-D bubble.
        /// </summary>
        public ExcelChartStyleEntry DataPoint
        {
            get
            {
                if (this._dataPoint == null)
                {
                    this._dataPoint = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataPoint", this);
                }
                return this._dataPoint;
            }
        }
        ExcelChartStyleEntry _dataPoint3D = null;
        /// <summary>
        /// Default formatting for a data point on a 3-D chart of type column, bar, line, pie, area and surface.
        /// </summary>
        public ExcelChartStyleEntry DataPoint3D
        {
            get
            {
                if (this._dataPoint3D == null)
                {
                    this._dataPoint3D = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataPoint3D", this);
                }
                return this._dataPoint3D;
            }
        }
        ExcelChartStyleEntry _dataPointLine = null;
        /// <summary>
        /// Default formatting for a data point on a 2-D chart of type line, scatter and radar
        /// </summary>
        public ExcelChartStyleEntry DataPointLine
        {
            get
            {
                if (this._dataPointLine == null)
                {
                    this._dataPointLine = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataPointLine", this);
                }
                return this._dataPointLine;
            }
        }
        ExcelChartStyleEntry _dataPointMarker = null;
        /// <summary>
        /// Default formatting for a datapoint marker
        /// </summary>
        public ExcelChartStyleEntry DataPointMarker
        {
            get
            {
                if (this._dataPointMarker == null)
                {
                    this._dataPointMarker = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataPointMarker", this);
                }
                return this._dataPointMarker;
            }
        }
        ExcelChartStyleMarkerLayout _dataPointMarkerLayout = null;
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
                    if(node == null)
                    {
                        throw new InvalidOperationException("Invalid Chartstyle xml: dataPointMarkerLayout element missing");
                    }

                    this._dataPointMarkerLayout = new ExcelChartStyleMarkerLayout(this.NameSpaceManager, node);
                }
                return this._dataPointMarkerLayout;
            }
        }
        ExcelChartStyleEntry _dataPointWireframe = null;
        /// <summary>
        /// Default formatting for a datapoint on a surface wireframe chart
        /// </summary>
        public ExcelChartStyleEntry DataPointWireframe
        {
            get
            {
                if (this._dataPointWireframe == null)
                {
                    this._dataPointWireframe = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataPointWireframe", this);
                }
                return this._dataPointWireframe;
            }
        }
        ExcelChartStyleEntry _dataTable = null;
        /// <summary>
        /// Default formatting for a Data table
        /// </summary>
        public ExcelChartStyleEntry DataTable
        {
            get
            {
                if (this._dataTable == null)
                {
                    this._dataTable = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dataTable", this);
                }
                return this._dataTable;
            }
        }
        ExcelChartStyleEntry _downBar = null;
        /// <summary>
        /// Default formatting for a downbar
        /// </summary>
        public ExcelChartStyleEntry DownBar
        {
            get
            {
                if (this._downBar == null)
                {
                    this._downBar = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:downBar", this);
                }
                return this._downBar;
            }
        }
        ExcelChartStyleEntry _dropLine = null;
        /// <summary>
        /// Default formatting for a dropline
        /// </summary>
        public ExcelChartStyleEntry DropLine
        {
            get
            {
                if (this._dropLine == null)
                {
                    this._dropLine = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:dropLine", this);
                }
                return this._dropLine;
            }
        }
        ExcelChartStyleEntry _errorBar = null;
        /// <summary>
        /// Default formatting for an errorbar
        /// </summary>
        public ExcelChartStyleEntry ErrorBar
        {
            get
            {
                if (this._errorBar == null)
                {
                    this._errorBar = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:errorBar", this);
                }
                return this._errorBar;
            }
        }
        ExcelChartStyleEntry _floor = null;
        /// <summary>
        /// Default formatting for a floor
        /// </summary>
        public ExcelChartStyleEntry Floor
        {
            get
            {
                if (this._floor == null)
                {
                    this._floor = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:floor", this);
                }
                return this._floor;
            }
        }
        ExcelChartStyleEntry _gridlineMajor = null;
        /// <summary>
        /// Default formatting for a major gridline
        /// </summary>
        public ExcelChartStyleEntry GridlineMajor
        {
            get
            {
                if (this._gridlineMajor == null)
                {
                    this._gridlineMajor = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:gridlineMajor", this);
                }
                return this._gridlineMajor;
            }
        }
        ExcelChartStyleEntry _gridlineMinor = null;
        /// <summary>
        /// Default formatting for a minor gridline
        /// </summary>
        public ExcelChartStyleEntry GridlineMinor
        {
            get
            {
                if (this._gridlineMinor == null)
                {
                    this._gridlineMinor = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:gridlineMinor", this);
                }
                return this._gridlineMinor;
            }
        }
        ExcelChartStyleEntry _hiLoLine = null;
        /// <summary>
        /// Default formatting for a high low line
        /// </summary>
        public ExcelChartStyleEntry HighLowLine
        {
            get
            {
                if (this._hiLoLine == null)
                {
                    this._hiLoLine = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:hiLoLine", this);
                }
                return this._hiLoLine;
            }
        }
        ExcelChartStyleEntry _leaderLine = null;
        /// <summary>
        /// Default formatting for a leader line
        /// </summary>
        public ExcelChartStyleEntry LeaderLine
        {
            get
            {
                if (this._leaderLine == null)
                {
                    this._leaderLine = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:leaderLine", this);
                }
                return this._leaderLine;
            }
        }
        /// <summary>
        /// Default formatting for a legend
        /// </summary>
        ExcelChartStyleEntry _legend = null;
        /// <summary>
        /// Default formatting for a chart legend
        /// </summary>
        public ExcelChartStyleEntry Legend
        {
            get
            {
                if (this._legend == null)
                {
                    this._legend = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:legend", this);
                }
                return this._legend;
            }
        }
        ExcelChartStyleEntry _plotArea = null;
        /// <summary>
        /// Default formatting for a plot area in a 2D chart
        /// </summary>
        public ExcelChartStyleEntry PlotArea
        {
            get
            {
                if (this._plotArea == null)
                {
                    this._plotArea = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:plotArea", this);
                }
                return this._plotArea;
            }
        }
        ExcelChartStyleEntry _plotArea3D = null;
        /// <summary>
        /// Default formatting for a plot area in a 3D chart
        /// </summary>
        public ExcelChartStyleEntry PlotArea3D
        {
            get
            {
                if (this._plotArea3D == null)
                {
                    this._plotArea3D = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:plotArea3D", this);
                }
                return this._plotArea3D;
            }
        }
        ExcelChartStyleEntry _seriesAxis = null;
        /// <summary>
        /// Default formatting for a series axis
        /// </summary>
        public ExcelChartStyleEntry SeriesAxis
        {
            get
            {
                if (this._seriesAxis == null)
                {
                    this._seriesAxis = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:seriesAxis", this);
                }
                return this._seriesAxis;
            }
        }
        ExcelChartStyleEntry _seriesLine = null;
        /// <summary>
        /// Default formatting for a series line
        /// </summary>
        public ExcelChartStyleEntry SeriesLine
        {
            get
            {
                if (this._seriesLine == null)
                {
                    this._seriesLine = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:seriesLine", this);
                }
                return this._seriesLine;
            }
        }
        ExcelChartStyleEntry _title = null;
        /// <summary>
        /// Default formatting for a chart title
        /// </summary>
        public ExcelChartStyleEntry Title
        {
            get
            {
                if (this._title == null)
                {
                    this._title = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:title", this);
                }
                return this._title;
            }
        }
        ExcelChartStyleEntry _trendline = null;
        /// <summary>
        /// Default formatting for a trend line
        /// </summary>
        public ExcelChartStyleEntry Trendline
        {
            get
            {
                if (this._trendline == null)
                {
                    this._trendline = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:trendline", this);
                }
                return this._trendline;
            }
        }
        ExcelChartStyleEntry _trendlineLabel = null;
        /// <summary>
        /// Default formatting for a trend line label
        /// </summary>
        public ExcelChartStyleEntry TrendlineLabel
        {
            get
            {
                if (this._trendlineLabel == null)
                {
                    this._trendlineLabel = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:trendlineLabel", this);
                }
                return this._trendlineLabel;
            }
        }
        ExcelChartStyleEntry _upBar = null;
        /// <summary>
        /// Default formatting for a up bar
        /// </summary>
        public ExcelChartStyleEntry UpBar
        {
            get
            {
                if (this._upBar == null)
                {
                    this._upBar = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:upBar", this);
                }
                return this._upBar;
            }
        }
        ExcelChartStyleEntry _valueAxis = null;
        /// <summary>
        /// Default formatting for a value axis
        /// </summary>
        public ExcelChartStyleEntry ValueAxis
        {
            get
            {
                if (this._valueAxis == null)
                {
                    this._valueAxis = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:valueAxis", this);
                }
                return this._valueAxis;
            }
        }
        ExcelChartStyleEntry _wall = null;
        /// <summary>
        /// Default formatting for a wall
        /// </summary>
        public ExcelChartStyleEntry Wall
        {
            get
            {
                if (this._wall == null)
                {
                    this._wall = new ExcelChartStyleEntry(this.NameSpaceManager, this.TopNode, "cs:wall", this);
                }
                return this._wall;
            }
        }

        /// <summary>
        /// The id of the chart style
        /// </summary>
        public int Id
        {
            get
            {
                return this.GetXmlNodeInt("@id");
            }
            internal set
            {
                this.SetXmlNodeString("@id", value.ToString(CultureInfo.InvariantCulture));
            }
        }

        ExcelPackage IPictureRelationDocument.Package => this._manager._chart._drawings._package;

        Dictionary<string, HashInfo> IPictureRelationDocument.Hashes => this._hashes;

        ZipPackagePart IPictureRelationDocument.RelatedPart => this._manager.StylePart;

        Uri IPictureRelationDocument.RelatedUri => this._manager.StyleUri;
    }
}
