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
using System.IO;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Drawing.Chart.ChartEx;
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Base class for Chart object.
    /// </summary>
    public abstract class ExcelChart : ExcelDrawing, IDrawingStyle, IStyleMandatoryProperties, IPictureRelationDocument
    {
        internal bool _isChartEx;
        internal const string topPath = "c:chartSpace";
        internal const string plotAreaPath = "c:chart/c:plotArea";
        //string _chartPath;
        internal ExcelChartAxis[] _axis;
        Dictionary<string, HashInfo> _hashes = new Dictionary<string, HashInfo>();
        /// <summary>
        /// The Xml helper for the chart xml
        /// </summary>
        protected internal XmlHelper _chartXmlHelper;
        internal ExcelChart _topChart = null;
        #region "Constructors"
        internal ExcelChart(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent, string drawingPath= "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
            base(drawings, node, drawingPath, nvPrPath, parent)
        {            
        }
        internal ExcelChart(ExcelDrawings drawings, XmlNode drawingsNode, XmlDocument chartXml = null, ExcelGroupShape parent=null, string drawingPath = "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
            base(drawings, drawingsNode, drawingPath, nvPrPath, parent)
        {
            this.Init(drawings, chartXml);
        }

        internal ExcelChart(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent, string drawingPath = "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
            base(topChart._drawings, topChart.TopNode, drawingPath, nvPrPath, parent)
        {
        }
        private void Init(ExcelDrawings drawings, XmlDocument chartXml)
        {
            this.WorkSheet = drawings.Worksheet;
            if (chartXml != null)
            {
                this.ChartXml = chartXml;
                this._chartXmlHelper = XmlHelperFactory.Create(drawings.NameSpaceManager, chartXml.DocumentElement);
            }
        }

        #endregion
        internal ExcelChartStyleManager _styleManager = null;
        /// <summary>
        /// Manage style settings for the chart
        /// </summary>
        public ExcelChartStyleManager StyleManager
        {
            get { return this._styleManager ??= new ExcelChartStyleManager(this.NameSpaceManager, this); }
        }
        private bool HasPrimaryAxis()
        {
            if (this._plotArea.ChartTypes.Count == 1)
            {
                return false;
            }
            foreach (ExcelChart? chart in this._plotArea.ChartTypes)
            {
                if (chart != this)
                {
                    if (chart.UseSecondaryAxis == false && chart.IsTypePieDoughnut() == false)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        internal abstract void AddAxis();
        bool _secondaryAxis = false;
        /// <summary>
        /// If true the charttype will use the secondary axis.
        /// The chart must contain a least one other charttype that uses the primary axis.
        /// </summary>
        public bool UseSecondaryAxis
        {
            get
            {
                return this._secondaryAxis;
            }
            set
            {
                if (this._secondaryAxis != value)
                {
                    if (value)
                    {
                        if (this.IsTypePieDoughnut())
                        {
                            throw (new Exception("Pie charts do not support axis"));
                        }
                        else if(this._isChartEx)
                        {
                            throw (new InvalidOperationException("Extentions charts don't support secondary axis"));
                        }
                        else if (this.HasPrimaryAxis() == false)
                        {
                            throw (new Exception("Can't set to secondary axis when no serie uses the primary axis"));
                        }
                        if (this.Axis.Length == 2)
                        {
                            this.AddAxis();
                        }
                        XmlNodeList? nl = this.ChartNode.SelectNodes("c:axId", this.NameSpaceManager);
                        nl[0].Attributes["val"].Value = this.Axis[2].Id;
                        nl[1].Attributes["val"].Value = this.Axis[3].Id;
                        this.XAxis = this.Axis[2];
                        this.YAxis = this.Axis[3];
                    }
                    else
                    {
                        XmlNodeList? nl = this.ChartNode.SelectNodes("c:axId", this.NameSpaceManager);
                        nl[0].Attributes["val"].Value = this.Axis[0].Id;
                        nl[1].Attributes["val"].Value = this.Axis[1].Id;
                        this.XAxis = this.Axis[0];
                        this.YAxis = this.Axis[1];
                    }

                    this._secondaryAxis = value;
                }
            }
        }
        #region "Properties"
        /// <summary>
        /// Reference to the worksheet
        /// </summary>
        public ExcelWorksheet WorkSheet { get; internal set; }
        /// <summary>
        /// The chart xml document
        /// </summary>
        public XmlDocument ChartXml { get; internal set; }
        /// <summary>
        /// The type of drawing
        /// </summary>
        public override eDrawingType DrawingType
        {
            get
            {
                return eDrawingType.Chart;
            }
        }

        /// <summary>
        /// Type of chart
        /// </summary>
        public eChartType ChartType { get; internal set; }
        /// <summary>
        /// The chart element
        /// </summary>
        internal protected XmlNode _chartNode = null;
        internal XmlNode ChartNode
        {
            get
            {
                return this._chartNode;
            }
        }
        internal ExcelChartTitle _title;
        /// <summary>
        /// The titel of the chart
        /// </summary>
        public virtual ExcelChartTitle Title
        {
            get { return this._title ??= this.GetTitle(); }
        }
        internal abstract ExcelChartTitle GetTitle();
        /// <summary>
        /// True if the chart has a title
        /// </summary>
        public abstract bool HasTitle
        {
            get;
        }
        /// <summary>
        /// If the chart has a legend
        /// </summary>
        public abstract bool HasLegend
        {
            get;
        }
        /// <summary>
        /// Remove the title from the chart
        /// </summary>
        public abstract void DeleteTitle();
        /// <summary>
        /// Chart series
        /// </summary>
        public virtual ExcelChartSeries<ExcelChartSerie> Series { get; } = new ExcelChartSeries<ExcelChartSerie>();
        /// <summary>
        /// An array containg all axis of all Charttypes
        /// </summary>
        public virtual ExcelChartAxis[] Axis
        {
            get
            {
                return this._axis;
            }
        }
        /// <summary>
        /// The X Axis
        /// </summary>
        public virtual ExcelChartAxis XAxis
        {
            get;
            internal protected set;
        }
        /// <summary>
        /// The Y Axis
        /// </summary>
        public virtual ExcelChartAxis YAxis
        {
            get;
            internal protected set;
        }
        /// <summary>
        /// The build-in chart styles. 
        /// Use <see cref="ExcelChart.StyleManager"/> for the more modern styling.
        /// </summary>
        public abstract eChartStyle Style
        {
            get;
            set;
        }
        internal ExcelChartPlotArea _plotArea = null;
        /// <summary>
        /// Plotarea
        /// </summary>
        public abstract ExcelChartPlotArea PlotArea
        {
            get;
        }
        internal ExcelChartLegend _legend = null;
        /// <summary>
        /// Legend
        /// </summary>
        public virtual ExcelChartLegend Legend
        {
            get
            {
                if (this._legend == null)
                {
                    if(this._isChartEx)
                    {
                        this._legend = new ExcelChartExLegend(this, this.NameSpaceManager, this.ChartXml.SelectSingleNode("cx:chartSpace/cx:chart/cx:legend", this.NameSpaceManager));
                    }
                    else
                    {
                        this._legend = new ExcelChartLegend(this.NameSpaceManager, this.ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", this.NameSpaceManager), this, "c");
                    }
                }
                return this._legend;
            }
        }
        /// <summary>
        /// Border
        /// </summary>
        public abstract ExcelDrawingBorder Border
        {
            get;
        }
        /// <summary>
        /// Access to Fill properties
        /// </summary>
        public abstract ExcelDrawingFill Fill
        {
            get;
        }
        /// <summary>
        /// Effects
        /// </summary>
        public abstract ExcelDrawingEffectStyle Effect
        {
            get;
        }
        /// <summary>
        /// 3D properties
        /// </summary>
        public abstract ExcelDrawing3D ThreeD
        {
            get;
        }
        /// <summary>
        /// Access to font properties
        /// </summary>
        public abstract ExcelTextFont Font
        {
            get;
        }
        /// <summary>
        /// Access to text body properties
        /// </summary>
        public abstract ExcelTextBody TextBody
        {
            get;
        }

        /// <summary>
        /// If the chart is a pivochart this is the pivotable used as source.
        /// </summary>
        public ExcelPivotTable PivotTableSource
        {
            get;
            protected set;
        }
        void IDrawingStyleBase.CreatespPr()
        {
            this._chartXmlHelper.CreatespPrNode("../../../c:spPr");
        }
        internal ZipPackagePart Part { get; set; }
        /// <summary>
        /// Package internal URI
        /// </summary>
        internal Uri UriChart { get; set; }
        internal new static string Id
        {
            get { return ""; }
        }
        #endregion
        #region "Chart type functions
        /// <summary>
        /// Returns true if the chart is a 3D chart
        /// </summary>
        /// <param name="chartType">The charttype to tests</param>
        /// <returns>True if the chart is a 3D chart</returns>
        internal static bool IsType3D(eChartType chartType)
        {
            return  chartType == eChartType.Area3D ||
                    chartType == eChartType.AreaStacked3D ||
                    chartType == eChartType.AreaStacked1003D ||
                    chartType == eChartType.BarClustered3D ||
                    chartType == eChartType.BarStacked3D ||
                    chartType == eChartType.BarStacked1003D ||
                    chartType == eChartType.Column3D ||
                    chartType == eChartType.ColumnClustered3D ||
                    chartType == eChartType.ColumnStacked3D ||
                    chartType == eChartType.ColumnStacked1003D ||
                    chartType == eChartType.Line3D ||
                    chartType == eChartType.Pie3D ||
                    chartType == eChartType.PieExploded3D ||
                    chartType == eChartType.ConeBarClustered ||
                    chartType == eChartType.ConeBarStacked ||
                    chartType == eChartType.ConeBarStacked100 ||
                    chartType == eChartType.ConeCol ||
                    chartType == eChartType.ConeColClustered ||
                    chartType == eChartType.ConeColStacked ||
                    chartType == eChartType.ConeColStacked100 ||
                    chartType == eChartType.CylinderBarClustered ||
                    chartType == eChartType.CylinderBarStacked ||
                    chartType == eChartType.CylinderBarStacked100 ||
                    chartType == eChartType.CylinderCol ||
                    chartType == eChartType.CylinderColClustered ||
                    chartType == eChartType.CylinderColStacked ||
                    chartType == eChartType.CylinderColStacked100 ||
                    chartType == eChartType.PyramidBarClustered ||
                    chartType == eChartType.PyramidBarStacked ||
                    chartType == eChartType.PyramidBarStacked100 ||
                    chartType == eChartType.PyramidCol ||
                    chartType == eChartType.PyramidColClustered ||
                    chartType == eChartType.PyramidColStacked ||
                    chartType == eChartType.PyramidColStacked100 ||
                    chartType == eChartType.Surface ||
                    chartType == eChartType.SurfaceTopView ||
                    chartType == eChartType.SurfaceTopViewWireframe ||
                    chartType == eChartType.SurfaceWireframe;
        }

        internal void ApplyStyleOnPart(IDrawingStyleBase chartPart, ExcelChartStyleEntry section, bool applyChartEx=false)
        {
            if((applyChartEx==false && this._isChartEx) || section == null)
            {
                return;
            }

            this._styleManager.ApplyStyle(chartPart, section);
        }

        /// <summary>
        /// Returns true if the chart is a 3D chart
        /// </summary>
        /// <returns>True if the chart is a 3D chart</returns>
        internal protected bool IsType3D()
        {
            return IsType3D(this.ChartType);
        }
        /// <summary>
        /// Returns true if the chart is a line chart
        /// </summary>
        /// <returns>True if the chart is a line chart</returns>
        protected internal bool IsTypeLine()
        {
            return this.ChartType == eChartType.Line || this.ChartType == eChartType.LineMarkers || this.ChartType == eChartType.LineMarkersStacked100 || this.ChartType == eChartType.LineStacked || this.ChartType == eChartType.LineStacked100 || this.ChartType == eChartType.Line3D;
        }
        /// <summary>
        /// Returns true if the chart is a radar chart
        /// </summary>
        /// <returns>True if the chart is a radar chart</returns>
        protected internal bool IsTypeRadar()
        {
            return this.ChartType == eChartType.Radar || this.ChartType == eChartType.RadarFilled || this.ChartType == eChartType.RadarMarkers;
        }

        /// <summary>
        /// Returns true if the chart is a scatter chart
        /// </summary>
        /// <returns>True if the chart is a scatter chart</returns>
        protected internal bool IsTypeScatter()
        {
            return this.ChartType == eChartType.XYScatter || this.ChartType == eChartType.XYScatterLines || this.ChartType == eChartType.XYScatterLinesNoMarkers || this.ChartType == eChartType.XYScatterSmooth || this.ChartType == eChartType.XYScatterSmoothNoMarkers;
        }
        /// <summary>
        /// Returns true if the chart is a bubble chart
        /// </summary>
        /// <returns>True if the chart is a bubble chart</returns>
        protected internal bool IsTypeBubble()
        {
            return this.ChartType == eChartType.Bubble || this.ChartType == eChartType.Bubble3DEffect;
        }
        /// <summary>
        /// Returns true if the chart is a scatter chart
        /// </summary>
        /// <returns>True if the chart is a scatter chart</returns>
        protected internal bool IsTypeArea()
        {
            return this.ChartType == eChartType.Area || this.ChartType == eChartType.AreaStacked || this.ChartType == eChartType.AreaStacked100 || this.ChartType == eChartType.AreaStacked1003D || this.ChartType == eChartType.AreaStacked3D;
        }

        /// <summary>
        /// Returns true if the chart is a sureface chart
        /// </summary>
        /// <returns>True if the chart is a sureface chart</returns>
        protected bool IsTypeSurface()
        {
            return this.ChartType == eChartType.Surface || this.ChartType == eChartType.SurfaceTopView || this.ChartType == eChartType.SurfaceTopViewWireframe || this.ChartType == eChartType.SurfaceWireframe;
        }
        /// <summary>
        /// Returns true if the chart is a sureface chart
        /// </summary>
        /// <returns>True if the chart is a sureface chart</returns>
        internal protected bool HasThirdAxis()
        {
            return this.IsTypeSurface() || this.ChartType == eChartType.Line3D || this.ChartType == eChartType.StockVHLC || this.ChartType == eChartType.StockVOHLC;
        }
        /// <summary>
        /// Returns true if the chart has shapes, like bars and columns
        /// </summary>
        /// <returns>True if the chart has shapes</returns>
        protected internal bool IsTypeShape()
        {
            return this.ChartType == eChartType.BarClustered3D || this.ChartType == eChartType.BarStacked3D || this.ChartType == eChartType.BarStacked1003D || this.ChartType == eChartType.BarClustered3D || this.ChartType == eChartType.BarStacked3D || this.ChartType == eChartType.BarStacked1003D || this.ChartType == eChartType.Column3D || this.ChartType == eChartType.ColumnClustered3D || this.ChartType == eChartType.ColumnStacked3D || this.ChartType == eChartType.ColumnStacked1003D ||
                   //ChartType == eChartType.3DPie ||
                   //ChartType == eChartType.3DPieExploded ||
                   //ChartType == eChartType.Bubble3DEffect ||
                   this.ChartType == eChartType.ConeBarClustered || this.ChartType == eChartType.ConeBarStacked || this.ChartType == eChartType.ConeBarStacked100 || this.ChartType == eChartType.ConeCol || this.ChartType == eChartType.ConeColClustered || this.ChartType == eChartType.ConeColStacked || this.ChartType == eChartType.ConeColStacked100 || this.ChartType == eChartType.CylinderBarClustered || this.ChartType == eChartType.CylinderBarStacked || this.ChartType == eChartType.CylinderBarStacked100 || this.ChartType == eChartType.CylinderCol || this.ChartType == eChartType.CylinderColClustered || this.ChartType == eChartType.CylinderColStacked || this.ChartType == eChartType.CylinderColStacked100 || this.ChartType == eChartType.PyramidBarClustered || this.ChartType == eChartType.PyramidBarStacked || this.ChartType == eChartType.PyramidBarStacked100 || this.ChartType == eChartType.PyramidCol || this.ChartType == eChartType.PyramidColClustered || this.ChartType == eChartType.PyramidColStacked || this.ChartType == eChartType.PyramidColStacked100; //||
                                                                  //ChartType == eChartType.Doughnut ||
                                                                  //ChartType == eChartType.DoughnutExploded;
        }
        /// <summary>
        /// Returns true if the chart is of type stacked percentage
        /// </summary>
        /// <returns>True if the chart is of type stacked percentage</returns>
        protected internal bool IsTypePercentStacked()
        {
            return this.ChartType == eChartType.AreaStacked100 || this.ChartType == eChartType.BarStacked100 || this.ChartType == eChartType.BarStacked1003D || this.ChartType == eChartType.ColumnStacked100 || this.ChartType == eChartType.ColumnStacked1003D || this.ChartType == eChartType.ConeBarStacked100 || this.ChartType == eChartType.ConeColStacked100 || this.ChartType == eChartType.CylinderBarStacked100 || this.ChartType == eChartType.CylinderColStacked || this.ChartType == eChartType.LineMarkersStacked100 || this.ChartType == eChartType.LineStacked100 || this.ChartType == eChartType.PyramidBarStacked100 || this.ChartType == eChartType.PyramidColStacked100;
        }
        /// <summary>
        /// Returns true if the chart is of type stacked 
        /// </summary>
        /// <returns>True if the chart is of type stacked</returns>
        protected internal bool IsTypeStacked()
        {
            return this.ChartType == eChartType.AreaStacked || this.ChartType == eChartType.AreaStacked3D || this.ChartType == eChartType.BarStacked || this.ChartType == eChartType.BarStacked3D || this.ChartType == eChartType.ColumnStacked3D || this.ChartType == eChartType.ColumnStacked || this.ChartType == eChartType.ConeBarStacked || this.ChartType == eChartType.ConeColStacked || this.ChartType == eChartType.CylinderBarStacked || this.ChartType == eChartType.CylinderColStacked || this.ChartType == eChartType.LineMarkersStacked || this.ChartType == eChartType.LineStacked || this.ChartType == eChartType.PyramidBarStacked || this.ChartType == eChartType.PyramidColStacked;
        }
        /// <summary>
        /// Returns true if the chart is of type clustered
        /// </summary>
        /// <returns>True if the chart is of type clustered</returns>
        protected bool IsTypeClustered()
        {
            return this.ChartType == eChartType.BarClustered || this.ChartType == eChartType.BarClustered3D || this.ChartType == eChartType.ColumnClustered3D || this.ChartType == eChartType.ColumnClustered || this.ChartType == eChartType.ConeBarClustered || this.ChartType == eChartType.ConeColClustered || this.ChartType == eChartType.CylinderBarClustered || this.ChartType == eChartType.CylinderColClustered || this.ChartType == eChartType.PyramidBarClustered || this.ChartType == eChartType.PyramidColClustered;
        }
        /// <summary>
        /// Returns true if the chart is a pie or Doughnut chart
        /// </summary>
        /// <returns>True if the chart is a pie or Doughnut chart</returns>
        protected internal bool IsTypePieDoughnut()
        {
            return this.IsTypePie() || this.IsTypeDoughnut();
        }
        /// <summary>
        /// Returns true if the chart is a Doughnut chart
        /// </summary>
        /// <returns>True if the chart is a Doughnut chart</returns>
        protected internal bool IsTypeDoughnut()
        {
            return this.ChartType == eChartType.Doughnut || this.ChartType == eChartType.DoughnutExploded;
        }
        /// <summary>
        /// Returns true if the chart is a pie chart
        /// </summary>
        /// <returns>true if the chart is a pie chart</returns>
        protected internal bool IsTypePie()
        {
            return this.ChartType == eChartType.Pie || this.ChartType == eChartType.PieExploded || this.ChartType == eChartType.PieOfPie || this.ChartType == eChartType.Pie3D || this.ChartType == eChartType.PieExploded3D || this.ChartType == eChartType.BarOfPie;
        }
        /// <summary>
        /// Return true if the chart is a stock chart.
        /// </summary>
        /// <returns>true if the chart is a stock chart.</returns>
        protected internal bool IsTypeStock()
        {
            return IsTypeStock(this.ChartType);
        }

        internal static bool IsTypeStock(eChartType chartType)
        {
            return chartType == eChartType.StockHLC ||
                   chartType == eChartType.StockOHLC ||
                   chartType == eChartType.StockVHLC ||
                   chartType == eChartType.StockVOHLC;
        }

        #endregion
        internal void InitChartTheme(int fallBackStyleId)
        {
            int styleId = fallBackStyleId + 100;
            XmlElement el = (XmlElement)this._chartXmlHelper.CreateNode("../../../mc:AlternateContent/mc:Choice");
            el.SetAttribute("xmlns:c14", ExcelPackage.schemaChart14);
            this._chartXmlHelper.SetXmlNodeString("../../../mc:AlternateContent/mc:Choice/@Requires", "c14");
            this._chartXmlHelper.SetXmlNodeString("../../../mc:AlternateContent/mc:Choice/c14:style/@val", styleId.ToString(CultureInfo.InvariantCulture));
            this._chartXmlHelper.SetXmlNodeString("../../../mc:AlternateContent/mc:Fallback/c:style/@val", fallBackStyleId.ToString(CultureInfo.InvariantCulture));
        }
        /// <summary>
        /// If the chart has only one serie this varies the colors for each point.
        /// This property does not apply to extention charts.
        /// </summary>
        public abstract bool VaryColors { get; set; }
        /// <summary>
        /// Formatting for the floor of a 3D chart. 
        /// <note type="note">This property is null for non 3D charts</note>
        /// </summary>
        public ExcelChartSurface Floor { get; protected set; } = null;
        /// <summary>
        /// Formatting for the sidewall of a 3D chart. 
        /// <note type="note">This property is null for non 3D charts</note>
        /// </summary>
        public ExcelChartSurface SideWall { get; protected set; } = null;
        /// <summary>
        /// Formatting for the backwall of a 3D chart. 
        /// <note type="note">This property is null for non 3D charts</note>
        /// </summary>
        public ExcelChartSurface BackWall { get; protected set; } = null; 
        internal override void DeleteMe()
        {
            try
            {
                this.Part.Package.DeletePart(this.UriChart);
            }
            catch (Exception ex)
            {
                throw (new InvalidDataException("EPPlus internal error when deleting chart.", ex));
            }
            base.DeleteMe();
        }
        /// <summary>
        /// Border rounded corners
        /// </summary>
        public abstract bool RoundedCorners
        {
            get;
            set;
        }
        /// <summary>
        /// Show data in hidden rows and columns
        /// </summary>
        public abstract bool ShowHiddenData
        {
            get;
            set;
        }
        /// <summary>
        /// Specifies the possible ways to display blanks
        /// </summary>
        public abstract eDisplayBlanksAs DisplayBlanksAs
        {
            get;
            set;
        }
        /// <summary>
        /// Specifies data labels over the maximum of the chart shall be shown
        /// </summary>
        public abstract bool ShowDataLabelsOverMaximum
        {
            get;
            set;
        }
        /// <summary>
        /// 3D-settings
        /// </summary>
        public abstract ExcelView3D View3D
        {
            get;
        }
        internal static ExcelChart GetChart(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null)
        {
            XmlNode chartNode;
            if (parent==null)
            {
                chartNode = node.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart", drawings.NameSpaceManager);
            }
            else
            {
                chartNode = node.SelectSingleNode("a:graphic/a:graphicData/c:chart", drawings.NameSpaceManager);
            }
            
            if (chartNode != null)
            {
                ZipPackageRelationship? drawingRelation = drawings.Part.GetRelationship(chartNode.Attributes["r:id"].Value);
                Uri? uriChart = UriHelper.ResolvePartUri(drawings.UriDrawing, drawingRelation.TargetUri);

                ZipPackagePart? part = drawings.Part.Package.GetPart(uriChart);
                XmlDocument? chartXml = new XmlDocument();
                LoadXmlSafe(chartXml, part.GetStream());

                return CreateChartFromXml(drawings, node, uriChart, part, chartXml, parent);
            }
            else
            {
                return null;
            }
        }
        internal static ExcelChartEx GetChartEx(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent = null)
        {
            XmlNode chartDrawingNode = node.SelectSingleNode("mc:AlternateContent/mc:Choice[@Requires='cx1' or @Requires='cx2']/xdr:graphicFrame/a:graphic/a:graphicData/cx:chart", drawings.NameSpaceManager);
            if (chartDrawingNode != null)
            {
                ZipPackageRelationship? drawingRelation = drawings.Part.GetRelationship(chartDrawingNode.Attributes["r:id"].Value);
                Uri? uriChart = UriHelper.ResolvePartUri(drawings.UriDrawing, drawingRelation.TargetUri);

                ZipPackagePart? part = drawings.Part.Package.GetPart(uriChart);
                XmlDocument? chartXml = new XmlDocument();
                LoadXmlSafe(chartXml, part.GetStream());

                XmlNode? chartNode = chartXml.SelectSingleNode("cx:chartSpace/cx:chart", drawings.NameSpaceManager);
                XmlNode? layoutId = chartNode.SelectSingleNode("cx:plotArea/cx:plotAreaRegion/cx:series[1]/@layoutId", drawings.NameSpaceManager);
                if(layoutId==null)
                {
                    return new ExcelTreemapChart(drawings, node, uriChart, part, chartXml, chartNode);
                }
                switch (layoutId.Value)
                {
                    case "treemap":
                        return new ExcelTreemapChart(drawings, node, uriChart, part, chartXml, chartNode);
                    case "sunburst":
                        return new ExcelSunburstChart(drawings, node, uriChart, part, chartXml, chartNode);
                    case "boxWhisker":
                        return new ExcelBoxWhiskerChart(drawings, node, uriChart, part, chartXml, chartNode);
                    case "clusteredColumn":
                    case "pareto":
                        return new ExcelHistogramChart(drawings, node, uriChart, part, chartXml, chartNode);
                    case "funnel":
                        return new ExcelFunnelChart(drawings, node, uriChart, part, chartXml, chartNode);
                    case "waterfall":
                        return new ExcelWaterfallChart(drawings, node, uriChart, part, chartXml, chartNode);
                    case "regionMap":
                        return new ExcelRegionMapChart(drawings, node, uriChart, part, chartXml, chartNode);
                    default:
                        throw new NotSupportedException($"Unsupported chart layout {layoutId.Value}");
                }
            }
            else
            {
                return null;
            }
        }
        internal static ExcelChart CreateChartFromXml(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, ExcelGroupShape parent = null)
        {
            ExcelChart topChart = null;
            foreach (XmlElement n in chartXml.SelectSingleNode(topPath + "/" + plotAreaPath, drawings.NameSpaceManager).ChildNodes)
            {
                if (n.LocalName.EndsWith("Chart"))
                {
                    if (topChart == null)
                    {
                        if (part == null)
                        {
                            topChart = GetChart(drawings, node);
                        }
                        else
                        {
                            topChart = GetChart(n, drawings, node, uriChart, part, chartXml, null, parent);
                        }
                    }
                    else
                    {
                        ExcelChart? subChart = GetChart(n, null, null, null, null, null, topChart, parent);
                        if (subChart != null)
                        {
                            topChart.PlotArea.ChartTypes.Add(subChart);                            
                        }
                    }
                }
            }
            return topChart;
        }
        internal static eChartType? GetChartTypeFromNodeName(string nodeName)
        {
            switch (nodeName)
            {
                case "stockChart":
                    return eChartType.StockHLC;
                case "area3DChart":
                case "areaChart":
                    return eChartType.Area;
                case "surface3DChart":
                case "surfaceChart":
                    return eChartType.Surface;
                case "radarChart":
                    return eChartType.Radar;
                case "bubbleChart":
                    return eChartType.Bubble;
                case "barChart":
                case "bar3DChart":
                    return eChartType.BarClustered;
                case "doughnutChart":
                    return eChartType.Doughnut;
                case "pie3DChart":
                case "pieChart":
                    return eChartType.Pie;
                case "ofPieChart":
                    return eChartType.PieOfPie;
                case "lineChart":
                case "line3DChart":
                    return eChartType.Line;
                case "scatterChart":
                    return eChartType.XYScatter;
                default:
                    return null;
            }
        }
        internal static ExcelChart GetNewChart(ExcelDrawings drawings, XmlNode drawNode, eChartType? chartType, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml = null)
        {
            switch (chartType)
            {
                case eChartType.Area:
                case eChartType.Area3D:
                case eChartType.AreaStacked:
                case eChartType.AreaStacked100:
                case eChartType.AreaStacked1003D:
                case eChartType.AreaStacked3D:
                    return new ExcelAreaChart(drawings, drawNode, chartType, topChart, PivotTableSource, chartXml);
                case eChartType.Pie:
                case eChartType.PieExploded:
                case eChartType.Pie3D:
                case eChartType.PieExploded3D:
                    return new ExcelPieChart(drawings, drawNode, chartType, topChart, PivotTableSource, chartXml);
                case eChartType.BarOfPie:
                case eChartType.PieOfPie:
                    return new ExcelOfPieChart(drawings, drawNode, chartType, topChart, PivotTableSource, chartXml);
                case eChartType.Doughnut:
                case eChartType.DoughnutExploded:
                    return new ExcelDoughnutChart(drawings, drawNode, chartType, topChart, PivotTableSource, chartXml);
                case eChartType.BarClustered:
                case eChartType.BarStacked:
                case eChartType.BarStacked100:
                case eChartType.BarClustered3D:
                case eChartType.BarStacked3D:
                case eChartType.BarStacked1003D:
                case eChartType.ConeBarClustered:
                case eChartType.ConeBarStacked:
                case eChartType.ConeBarStacked100:
                case eChartType.CylinderBarClustered:
                case eChartType.CylinderBarStacked:
                case eChartType.CylinderBarStacked100:
                case eChartType.PyramidBarClustered:
                case eChartType.PyramidBarStacked:
                case eChartType.PyramidBarStacked100:
                case eChartType.ColumnClustered:
                case eChartType.ColumnStacked:
                case eChartType.ColumnStacked100:
                case eChartType.Column3D:
                case eChartType.ColumnClustered3D:
                case eChartType.ColumnStacked3D:
                case eChartType.ColumnStacked1003D:
                case eChartType.ConeCol:
                case eChartType.ConeColClustered:
                case eChartType.ConeColStacked:
                case eChartType.ConeColStacked100:
                case eChartType.CylinderCol:
                case eChartType.CylinderColClustered:
                case eChartType.CylinderColStacked:
                case eChartType.CylinderColStacked100:
                case eChartType.PyramidCol:
                case eChartType.PyramidColClustered:
                case eChartType.PyramidColStacked:
                case eChartType.PyramidColStacked100:
                    return new ExcelBarChart(drawings, drawNode, chartType, topChart, PivotTableSource, chartXml);
                case eChartType.XYScatter:
                case eChartType.XYScatterLines:
                case eChartType.XYScatterLinesNoMarkers:
                case eChartType.XYScatterSmooth:
                case eChartType.XYScatterSmoothNoMarkers:
                    return new ExcelScatterChart(drawings, drawNode, chartType, topChart, PivotTableSource, chartXml);
                case eChartType.Line:
                case eChartType.Line3D:
                case eChartType.LineMarkers:
                case eChartType.LineMarkersStacked:
                case eChartType.LineMarkersStacked100:
                case eChartType.LineStacked:
                case eChartType.LineStacked100:
                    return new ExcelLineChart(drawings, drawNode, chartType, topChart, PivotTableSource, chartXml);
                case eChartType.Bubble:
                case eChartType.Bubble3DEffect:
                    return new ExcelBubbleChart(drawings, drawNode, chartType, topChart, PivotTableSource, chartXml);
                case eChartType.Radar:
                case eChartType.RadarFilled:
                case eChartType.RadarMarkers:
                    return new ExcelRadarChart(drawings, drawNode, chartType, topChart, PivotTableSource, chartXml);
                case eChartType.Surface:
                case eChartType.SurfaceTopView:
                case eChartType.SurfaceTopViewWireframe:
                case eChartType.SurfaceWireframe:
                    return new ExcelSurfaceChart(drawings, drawNode, chartType, topChart, PivotTableSource, chartXml);
                case eChartType.StockHLC:
                case eChartType.StockOHLC:
                case eChartType.StockVHLC:
                case eChartType.StockVOHLC:
                    return new ExcelStockChart(drawings, drawNode, chartType, topChart, PivotTableSource, chartXml);
                case eChartType.Sunburst:
                    return new ExcelSunburstChart(drawings, drawNode, chartType, chartXml);
                case eChartType.Treemap:
                    return new ExcelTreemapChart(drawings, drawNode, chartType, chartXml);
                case eChartType.BoxWhisker:
                    return new ExcelBoxWhiskerChart(drawings, drawNode, chartType, chartXml);
                case eChartType.Histogram:
                case eChartType.Pareto:
                    return new ExcelHistogramChart(drawings, drawNode, chartType, chartXml);
                case eChartType.Waterfall:
                    return new ExcelWaterfallChart(drawings, drawNode, chartType, chartXml);
                case eChartType.Funnel:
                    return new ExcelFunnelChart(drawings, drawNode, chartType, chartXml);
                case eChartType.RegionMap:
                    return new ExcelRegionMapChart(drawings, drawNode, chartType, chartXml);
                default:
                    return new ExcelChartStandard(drawings, drawNode, chartType, topChart, PivotTableSource, chartXml);

            }
        }
        internal static ExcelChart GetChart(XmlElement chartNode, ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, ExcelChart topChart, ExcelGroupShape parent)
        {
            switch (chartNode.LocalName)
            {
                case "stockChart":
                    if (topChart == null)
                    {
                        return new ExcelStockChart(drawings, node, uriChart, part, chartXml, chartNode, parent);
                    }
                    else
                    {
                        if(topChart is ExcelStockChart chart)
                        {
                            return chart;
                        }
                        else
                        {
                            return new ExcelStockChart(topChart, chartNode, parent);
                        }
                    }
                case "area3DChart":
                case "areaChart":
                    if (topChart == null)
                    {
                        return new ExcelAreaChart(drawings, node, uriChart, part, chartXml, chartNode, parent);
                    }
                    else
                    {
                        return new ExcelAreaChart(topChart, chartNode, parent);
                    }
                case "surface3DChart":
                case "surfaceChart":
                    if (topChart == null)
                    {
                        return new ExcelSurfaceChart(drawings, node, uriChart, part, chartXml, chartNode, parent);
                    }
                    else
                    {
                        return new ExcelSurfaceChart(topChart, chartNode, parent);
                    }
                case "radarChart":
                    if (topChart == null)
                    {
                        return new ExcelRadarChart(drawings, node, uriChart, part, chartXml, chartNode, parent);
                    }
                    else
                    {
                        return new ExcelRadarChart(topChart, chartNode, parent);
                    }
                case "bubbleChart":
                    if (topChart == null)
                    {
                        return new ExcelBubbleChart(drawings, node, uriChart, part, chartXml, chartNode, parent);
                    }
                    else
                    {
                        return new ExcelBubbleChart(topChart, chartNode, parent);
                    }
                case "barChart":
                case "bar3DChart":
                    if (topChart == null)
                    {
                        if (chartNode.LocalName == "barChart" && chartNode.NextSibling?.LocalName == "stockChart")
                        {
                            return new ExcelStockChart(drawings, node, uriChart, part, chartXml, chartNode, parent);
                        }
                        else
                        {
                            return new ExcelBarChart(drawings, node, uriChart, part, chartXml, chartNode, parent);
                        }
                    }
                    else
                    {
                        return new ExcelBarChart(topChart, chartNode, parent);
                    }
                case "doughnutChart":
                    if (topChart == null)
                    {
                        return new ExcelDoughnutChart(drawings, node, uriChart, part, chartXml, chartNode, parent);
                    }
                    else
                    {
                        return new ExcelDoughnutChart(topChart, chartNode, parent);
                    }
                case "pie3DChart":
                case "pieChart":
                    if (topChart == null)
                    {
                        return new ExcelPieChart(drawings, node, uriChart, part, chartXml, chartNode, parent);
                    }
                    else
                    {
                        return new ExcelPieChart(topChart, chartNode, parent);
                    }
                case "ofPieChart":
                    if (topChart == null)
                    {
                        return new ExcelOfPieChart(drawings, node, uriChart, part, chartXml, chartNode, parent);
                    }
                    else
                    {
                        return new ExcelOfPieChart(topChart, chartNode, parent);
                    }
                case "lineChart":
                case "line3DChart":
                    if (topChart == null)
                    {
                        if (uriChart == null)
                        {
                            return new ExcelLineChart(drawings, node, eChartType.Line, null, null, chartXml, parent);
                        }
                        else
                        {
                            return new ExcelLineChart(drawings, node, uriChart, part, chartXml, chartNode, parent);
                        }
                    }
                    else
                    {
                        return new ExcelLineChart(topChart, chartNode, parent);
                    }
                case "scatterChart":
                    if (topChart == null)
                    {
                        return new ExcelScatterChart(drawings, node, uriChart, part, chartXml, chartNode, parent);
                    }
                    else
                    {
                        return new ExcelScatterChart(topChart, chartNode, parent);
                    }
                default:
                    return null;
            }
        }
        void IStyleMandatoryProperties.SetMandatoryProperties()
        {
            this._chartXmlHelper.CreatespPrNode("../c:spPr");
        }
        ExcelPackage IPictureRelationDocument.Package => this._drawings._package;

        Dictionary<string, HashInfo> IPictureRelationDocument.Hashes => this._hashes;

        ZipPackagePart IPictureRelationDocument.RelatedPart => this.Part;

        Uri IPictureRelationDocument.RelatedUri => this.UriChart;
    }
}
