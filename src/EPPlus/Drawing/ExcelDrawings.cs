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

using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml.Table;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System.Xml.Linq;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif
using System.Xml;

namespace OfficeOpenXml.Drawing;

/// <summary>
/// Collection for Drawing objects.
/// </summary>
public class ExcelDrawings : IEnumerable<ExcelDrawing>, IDisposable, IPictureRelationDocument
{
    private XmlDocument _drawingsXml = new XmlDocument();
    internal Dictionary<string, int> _drawingNames;
    internal List<ExcelDrawing> _drawingsList;
    Dictionary<string, HashInfo> _hashes = new Dictionary<string, HashInfo>();

    internal class ImageCompare
    {
        internal byte[] image { get; set; }

        internal string relID { get; set; }

        internal bool Comparer(byte[] compareImg)
        {
            if (compareImg.Length != this.image.Length)
            {
                return false;
            }

            for (int i = 0; i < this.image.Length; i++)
            {
                if (this.image[i] != compareImg[i])
                {
                    return false;
                }
            }

            return true; //Equal
        }
    }

    internal ExcelPackage _package;
    internal ZipPackageRelationship _drawingRelation = null;
    internal string _seriesTemplateXml;

    internal ExcelDrawings(ExcelPackage xlPackage, ExcelWorksheet sheet)
    {
        xlPackage.Workbook.LoadAllDrawings(sheet.Name);

        this._package = xlPackage;
        this.Worksheet = sheet;

        this._drawingsXml = new XmlDocument();
        this._drawingsXml.PreserveWhitespace = true;
        this._drawingsList = new List<ExcelDrawing>();
        this._drawingNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        this.CreateNSM();
        XmlNode node = sheet.WorksheetXml.SelectSingleNode("//d:drawing", sheet.NameSpaceManager);

        if (node != null && sheet != null)
        {
            this._drawingRelation = sheet.Part.GetRelationship(node.Attributes["r:id"].Value);
            this._uriDrawing = UriHelper.ResolvePartUri(sheet.WorksheetUri, this._drawingRelation.TargetUri);

            this._part = xlPackage.ZipPackage.GetPart(this._uriDrawing);
            XmlHelper.LoadXmlSafe(this._drawingsXml, this._part.GetStream());

            this.AddDrawings();
        }
    }

    internal ExcelWorksheet Worksheet { get; set; }

    /// <summary>
    /// A reference to the drawing xml document
    /// </summary>
    public XmlDocument DrawingXml
    {
        get { return this._drawingsXml; }
    }

    private void AddDrawings()
    {
        XmlNodeList list =
            this._drawingsXml.SelectNodes("//*[self::xdr:oneCellAnchor or self::xdr:twoCellAnchor or self::xdr:absoluteAnchor]", this.NameSpaceManager);

        foreach (XmlNode node in list)
        {
            ExcelDrawing dr;

            switch (node.LocalName)
            {
                case "oneCellAnchor":
                case "twoCellAnchor":
                case "absoluteAnchor":
                    dr = ExcelDrawing.GetDrawing(this, node);

                    break;

                default:
                    dr = null;

                    break;
            }

            if (dr != null)
            {
                this._drawingsList.Add(dr);

                if (!this._drawingNames.ContainsKey(dr.Name))
                {
                    this._drawingNames.Add(dr.Name, this._drawingsList.Count - 1);
                }
            }
        }
    }

    #region NamespaceManager

    /// <summary>
    /// Creates the NamespaceManager. 
    /// </summary>
    private void CreateNSM()
    {
        NameTable nt = new NameTable();
        this.NameSpaceManager = new XmlNamespaceManager(nt);
        this.NameSpaceManager.AddNamespace("d", ExcelPackage.schemaMain);
        this.NameSpaceManager.AddNamespace("a", ExcelPackage.schemaDrawings);
        this.NameSpaceManager.AddNamespace("xdr", ExcelPackage.schemaSheetDrawings);
        this.NameSpaceManager.AddNamespace("c", ExcelPackage.schemaChart);
        this.NameSpaceManager.AddNamespace("r", ExcelPackage.schemaRelationships);
        this.NameSpaceManager.AddNamespace("cs", ExcelPackage.schemaChartStyle);
        this.NameSpaceManager.AddNamespace("mc", ExcelPackage.schemaMarkupCompatibility);
        this.NameSpaceManager.AddNamespace("c14", ExcelPackage.schemaChart14);
        this.NameSpaceManager.AddNamespace("mc", ExcelPackage.schemaMc2006);
        this.NameSpaceManager.AddNamespace("cx", ExcelPackage.schemaChartExMain);
        this.NameSpaceManager.AddNamespace("cx1", ExcelPackage.schemaChartEx2015_9_8);
        this.NameSpaceManager.AddNamespace("cx2", ExcelPackage.schemaChartEx2015_10_21);
        this.NameSpaceManager.AddNamespace("x14", ExcelPackage.schemaMainX14);
        this.NameSpaceManager.AddNamespace("x15", ExcelPackage.schemaMainX15);
        this.NameSpaceManager.AddNamespace("sle", ExcelPackage.schemaSlicer2010);
        this.NameSpaceManager.AddNamespace("sle15", ExcelPackage.schemaSlicer);
        this.NameSpaceManager.AddNamespace("a14", ExcelPackage.schemaDrawings2010);
        this.NameSpaceManager.AddNamespace("asvg", "http://schemas.microsoft.com/office/drawing/2016/SVG/main");
    }

    internal XmlNamespaceManager NameSpaceManager { get; private set; } = null;

    #endregion

    #region IEnumerable Members

    /// <summary>
    /// Get the enumerator
    /// </summary>
    /// <returns>The enumerator</returns>
    IEnumerator IEnumerable.GetEnumerator()
    {
        return this._drawingsList.GetEnumerator();
    }

    #region IEnumerable<ExcelDrawing> Members

    IEnumerator<ExcelDrawing> IEnumerable<ExcelDrawing>.GetEnumerator()
    {
        return this._drawingsList.GetEnumerator();
    }

    #endregion

    /// <summary>
    /// Returns the drawing at the specified position.  
    /// </summary>
    /// <param name="PositionID">The position of the drawing. 0-base</param>
    /// <returns></returns>
    public ExcelDrawing this[int PositionID]
    {
        get { return this._drawingsList[PositionID]; }
    }

    /// <summary>
    /// Returns the drawing matching the specified name
    /// </summary>
    /// <param name="Name">The name of the worksheet</param>
    /// <returns></returns>
    public ExcelDrawing this[string Name]
    {
        get
        {
            if (this._drawingNames.ContainsKey(Name))
            {
                return this._drawingsList[this._drawingNames[Name]];
            }
            else
            {
                return null;
            }
        }
    }

    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count
    {
        get
        {
            if (this._drawingsList == null)
            {
                return 0;
            }
            else
            {
                return this._drawingsList.Count;
            }
        }
    }

    ZipPackagePart _part = null;

    internal ZipPackagePart Part
    {
        get { return this._part; }
    }

    Uri _uriDrawing = null;
    internal int _nextChartStyleId = 100;

    /// <summary>
    /// The uri to the drawing xml file inside the package
    /// </summary>
    internal Uri UriDrawing
    {
        get { return this._uriDrawing; }
    }

    ExcelPackage IPictureRelationDocument.Package => this._package;

    Dictionary<string, HashInfo> IPictureRelationDocument.Hashes => this._hashes;

    ZipPackagePart IPictureRelationDocument.RelatedPart => this._part;

    Uri IPictureRelationDocument.RelatedUri => this._uriDrawing;

    #endregion

    #region Add functions

    /// <summary>
    /// Adds a new chart to the worksheet.
    /// Stock charts cannot be added by this method. See <see cref="ExcelDrawings.AddStockChart(string, eStockChartType, ExcelRangeBase)"/>
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <param name="PivotTableSource">The pivottable source for a pivotchart</param>
    /// <param name="DrawingType">The top element drawingtype. Default is OneCellAnchor for Pictures and TwoCellAnchor from Charts and Shapes</param>
    /// <returns>The chart</returns>
    public ExcelChart AddChart(string Name, eChartType ChartType, ExcelPivotTable PivotTableSource, eEditAs DrawingType = eEditAs.TwoCell)
    {
        if (ExcelChart.IsTypeStock(ChartType))
        {
            throw new InvalidOperationException("For stock charts please use the AddStockChart method.");
        }

        return this.AddAllChartTypes(Name, ChartType, PivotTableSource, DrawingType);
    }

    internal ExcelChart AddAllChartTypes(string Name, eChartType ChartType, ExcelPivotTable PivotTableSource, eEditAs DrawingType = eEditAs.TwoCell)
    {
        if (this._drawingNames.ContainsKey(Name))
        {
            throw new Exception("Name already exists in the drawings collection");
        }

        if (this.Worksheet is ExcelChartsheet && this._drawingsList.Count > 0)
        {
            throw new InvalidOperationException("Chart Worksheets can't have more than one chart");
        }

        XmlElement drawNode = this.CreateDrawingXml(DrawingType);

        ExcelChart? chart = ExcelChart.GetNewChart(this, drawNode, ChartType, null, PivotTableSource);
        chart.Name = Name;
        this._drawingsList.Add(chart);
        this._drawingNames.Add(Name, this._drawingsList.Count - 1);

        return chart;
    }

    /// <summary>
    /// Adds a new chart to the worksheet.
    /// Do not support Stock charts . 
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <returns>The chart</returns>
    public ExcelChart AddChart(string Name, eChartType ChartType)
    {
        return this.AddChart(Name, ChartType, null);
    }

    /// <summary>
    /// Adds a new chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <returns>The chart</returns>
    public ExcelChartEx AddExtendedChart(string Name, eChartExType ChartType)
    {
        return (ExcelChartEx)this.AddAllChartTypes(Name, (eChartType)ChartType, null);
    }

    /// <summary>
    /// Adds a new sunburst chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <returns>The chart</returns>
    public ExcelSunburstChart AddSunburstChart(string Name)
    {
        return (ExcelSunburstChart)this.AddAllChartTypes(Name, eChartType.Sunburst, null);
    }

    /// <summary>
    /// Adds a new treemap chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <returns>The chart</returns>
    public ExcelTreemapChart AddTreemapChart(string Name)
    {
        return (ExcelTreemapChart)this.AddAllChartTypes(Name, eChartType.Treemap, null);
    }

    /// <summary>
    /// Adds a new box &amp; whisker chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <returns>The chart</returns>
    public ExcelBoxWhiskerChart AddBoxWhiskerChart(string Name)
    {
        return (ExcelBoxWhiskerChart)this.AddAllChartTypes(Name, eChartType.BoxWhisker, null);
    }

    /// <summary>
    /// Adds a new Histogram or Pareto chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="AddParetoLine">If true a pareto line is added to the chart. The <see cref="ExcelChart.ChartType"/> will also be Pareto.</param>
    /// <returns>The chart</returns>
    public ExcelHistogramChart AddHistogramChart(string Name, bool AddParetoLine = false)
    {
        return (ExcelHistogramChart)this.AddAllChartTypes(Name, AddParetoLine ? eChartType.Pareto : eChartType.Histogram, null);
    }

    /// <summary>
    /// Adds a waterfall chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <returns>The chart</returns>
    public ExcelWaterfallChart AddWaterfallChart(string Name)
    {
        return (ExcelWaterfallChart)this.AddAllChartTypes(Name, eChartType.Waterfall, null);
    }

    /// <summary>
    /// Adds a funnel chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <returns>The chart</returns>
    public ExcelFunnelChart AddFunnelChart(string Name)
    {
        return (ExcelFunnelChart)this.AddAllChartTypes(Name, eChartType.Funnel, null);
    }

    /// <summary>
    /// Adds a region map chart to the worksheet.
    /// Note that EPPlus rely on the spreadsheet application to create the geocache data
    /// </summary>
    /// <param name="Name"></param>
    /// <returns>The chart</returns>
    public ExcelRegionMapChart AddRegionMapChart(string Name)
    {
        return (ExcelRegionMapChart)this.AddAllChartTypes(Name, eChartType.RegionMap, null);
    }

    /// <summary>
    /// Adds a new extended chart to the worksheet.
    /// Extended charts are 
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <param name="PivotTableSource">The pivottable source for a pivotchart</param>    
    /// <returns>The chart</returns>
    public ExcelChartEx AddExtendedChart(string Name, eChartExType ChartType, ExcelPivotTable PivotTableSource)
    {
        return (ExcelChartEx)this.AddAllChartTypes(Name, (eChartType)ChartType, PivotTableSource);
    }

    /// <summary>
    /// Adds a new stock chart to the worksheet.
    /// Requires a range with four, five or six columns depending on the stock chart type.
    /// The first column is the category series. 
    /// The following columns in the range depend on the stock chart type (HLC, OHLC, VHLC, VOHLC).
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">The Stock chart type</param>
    /// <param name="Range">The category serie. A serie containng dates </param>
    /// <returns>The chart</returns>
    public ExcelStockChart AddStockChart(string Name, eStockChartType ChartType, ExcelRangeBase Range)
    {
        int startRow = Range.Start.Row;
        int startCol = Range.Start.Column;
        int endRow = Range.End.Row;
        ExcelWorksheet? ws = Range.Worksheet;

        switch (ChartType)
        {
            case eStockChartType.StockHLC:
                if (Range.Columns != 4)
                {
                    throw new
                        InvalidOperationException("Range must contain 4 columns with the Category serie to the left and the High Price, Low Price and Close Price series");
                }

                return this.AddStockChart(Name,
                                          ws.Cells[startRow, startCol, endRow, startCol],
                                          ws.Cells[startRow, startCol + 1, endRow, startCol + 1],
                                          ws.Cells[startRow, startCol + 2, endRow, startCol + 2],
                                          ws.Cells[startRow, startCol + 3, endRow, startCol + 3]);

            case eStockChartType.StockOHLC:
                if (Range.Columns != 5)
                {
                    throw new
                        InvalidOperationException("Range must contain 5 columns with the Category serie to the left and the Opening Price, High Price, Low Price and Close Price series");
                }

                return this.AddStockChart(Name,
                                          ws.Cells[startRow, startCol, endRow, startCol],
                                          ws.Cells[startRow, startCol + 2, endRow, startCol + 2],
                                          ws.Cells[startRow, startCol + 3, endRow, startCol + 3],
                                          ws.Cells[startRow, startCol + 4, endRow, startCol + 4],
                                          ws.Cells[startRow, startCol + 1, endRow, startCol + 1]);

            case eStockChartType.StockVHLC:
                if (Range.Columns != 5)
                {
                    throw new
                        InvalidOperationException("Range must contain 5 columns with the Category serie to the left and the Volume, High Price, Low Price and Close Price series");
                }

                return this.AddStockChart(Name,
                                          ws.Cells[startRow, startCol, endRow, startCol],
                                          ws.Cells[startRow, startCol + 2, endRow, startCol + 2],
                                          ws.Cells[startRow, startCol + 3, endRow, startCol + 3],
                                          ws.Cells[startRow, startCol + 4, endRow, startCol + 4],
                                          null,
                                          ws.Cells[startRow, startCol + 1, endRow, startCol + 1]);

            case eStockChartType.StockVOHLC:
                if (Range.Columns != 6)
                {
                    throw new
                        InvalidOperationException("Range must contain 6 columns with the Category serie to the left and the Volume, Opening Price, High Price, Low Price and Close Price series");
                }

                return this.AddStockChart(Name,
                                          ws.Cells[startRow, startCol, endRow, startCol],
                                          ws.Cells[startRow, startCol + 3, endRow, startCol + 3],
                                          ws.Cells[startRow, startCol + 4, endRow, startCol + 4],
                                          ws.Cells[startRow, startCol + 5, endRow, startCol + 5],
                                          ws.Cells[startRow, startCol + 2, endRow, startCol + 2],
                                          ws.Cells[startRow, startCol + 1, endRow, startCol + 1]);

            default:
                throw new InvalidOperationException("Unknown eStockChartType");
        }
    }

    /// <summary>
    /// Adds a new stock chart to the worksheet.
    /// The stock chart type will depend on if the parameters OpenSerie and/or VolumeSerie is supplied
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="CategorySerie">The category serie. A serie containng dates </param>
    /// <param name="HighSerie">The high price serie</param>    
    /// <param name="LowSerie">The low price serie</param>    
    /// <param name="CloseSerie">The close price serie containing</param>    
    /// <param name="OpenSerie">The opening price serie. Supplying this serie will create a StockOHLC or StockVOHLC chart</param>
    /// <param name="VolumeSerie">The volume represented as a column chart. Supplying this serie will create a StockVHLC or StockVOHLC chart</param>
    /// <returns>The chart</returns>
    public ExcelStockChart AddStockChart(string Name,
                                         ExcelRangeBase CategorySerie,
                                         ExcelRangeBase HighSerie,
                                         ExcelRangeBase LowSerie,
                                         ExcelRangeBase CloseSerie,
                                         ExcelRangeBase OpenSerie = null,
                                         ExcelRangeBase VolumeSerie = null)
    {
        ValidateSeries(CategorySerie, LowSerie, HighSerie, CloseSerie);

        eChartType chartType = ExcelStockChart.GetChartType(OpenSerie, VolumeSerie);

        ExcelStockChart? chart = (ExcelStockChart)this.AddAllChartTypes(Name, chartType, null);

        if (CategorySerie.Rows > 1)
        {
            if (CategorySerie.Offset(1, 0, 1, 1).Value is string)
            {
                chart.XAxis.ChangeAxisType(eAxisType.Date);
            }
        }

        _ = chart.AddHighLowLines();

        if (chartType == eChartType.StockOHLC || chartType == eChartType.StockVOHLC)
        {
            chart.AddUpDownBars(true, true);
        }

        if (chartType == eChartType.StockVHLC || chartType == eChartType.StockVOHLC)
        {
            _ = chart.PlotArea.ChartTypes[0].Series.Add(VolumeSerie, CategorySerie);
        }

        if (chartType == eChartType.StockOHLC || chartType == eChartType.StockVOHLC)
        {
            _ = chart.Series.Add(OpenSerie, CategorySerie);
        }

        _ = chart.Series.Add(HighSerie, CategorySerie);
        _ = chart.Series.Add(LowSerie, CategorySerie);
        _ = chart.Series.Add(CloseSerie, CategorySerie);

        return chart;
    }

    /// <summary>
    /// Adds a new stock chart to the worksheet.
    /// The stock chart type will depend on if the parameters OpenSerie and/or VolumeSerie is supplied
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="CategorySerie">The category serie. A serie containing dates </param>
    /// <param name="HighSerie">The high price serie</param>    
    /// <param name="LowSerie">The low price serie</param>    
    /// <param name="CloseSerie">The close price serie containing</param>    
    /// <param name="OpenSerie">The opening price serie. Supplying this serie will create a StockOHLC or StockVOHLC chart</param>
    /// <param name="VolumeSerie">The volume represented as a column chart. Supplying this serie will create a StockVHLC or StockVOHLC chart</param>
    /// <returns>The chart</returns>
    public ExcelStockChart AddStockChart(string Name,
                                         string CategorySerie,
                                         string HighSerie,
                                         string LowSerie,
                                         string CloseSerie,
                                         string OpenSerie = null,
                                         string VolumeSerie = null)
    {
        eChartType chartType = ExcelStockChart.GetChartType(OpenSerie, VolumeSerie);

        ExcelStockChart? chart = (ExcelStockChart)this.AddAllChartTypes(Name, chartType, null);
        ExcelStockChart.SetStockChartSeries(chart, chartType, CategorySerie, HighSerie, LowSerie, CloseSerie, OpenSerie, VolumeSerie);

        return chart;
    }

    private static void ValidateSeries(ExcelRangeBase CategorySerie, ExcelRangeBase HighSerie, ExcelRangeBase LowSerie, ExcelRangeBase CloseSerie)
    {
        if (CategorySerie == null)
        {
            throw new ArgumentNullException("CategorySerie");
        }
        else if (HighSerie == null)
        {
            throw new ArgumentNullException("HighSerie");
        }
        else if (LowSerie == null)
        {
            throw new ArgumentNullException("LowSerie");
        }
        else if (CloseSerie == null)
        {
            throw new ArgumentNullException("CloseSerie ");
        }
    }

    /// <summary>
    /// Add a new linechart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of linechart</param>
    /// <returns>The chart</returns>
    public ExcelLineChart AddLineChart(string Name, eLineChartType ChartType)
    {
        return (ExcelLineChart)this.AddAllChartTypes(Name, (eChartType)ChartType, null);
    }

    /// <summary>
    /// Adds a new linechart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <param name="PivotTableSource">The pivottable source for a pivotchart</param>    
    /// <returns>The chart</returns>
    public ExcelLineChart AddLineChart(string Name, eLineChartType ChartType, ExcelPivotTable PivotTableSource)
    {
        return (ExcelLineChart)this.AddAllChartTypes(Name, (eChartType)ChartType, PivotTableSource);
    }

    /// <summary>
    /// Add a new area chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of linechart</param>
    /// <returns>The chart</returns>
    public ExcelAreaChart AddAreaChart(string Name, eAreaChartType ChartType)
    {
        return (ExcelAreaChart)this.AddAllChartTypes(Name, (eChartType)ChartType, null);
    }

    /// <summary>
    /// Adds a new area chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <param name="PivotTableSource">The pivottable source for a pivotchart</param>    
    /// <returns>The chart</returns>
    public ExcelAreaChart AddAreaChart(string Name, eAreaChartType ChartType, ExcelPivotTable PivotTableSource)
    {
        return (ExcelAreaChart)this.AddAllChartTypes(Name, (eChartType)ChartType, PivotTableSource);
    }

    /// <summary>
    /// Adds a new barchart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of linechart</param>
    /// <returns>The chart</returns>
    public ExcelBarChart AddBarChart(string Name, eBarChartType ChartType)
    {
        return (ExcelBarChart)this.AddAllChartTypes(Name, (eChartType)ChartType, null);
    }

    /// <summary>
    /// Adds a new column- or bar- chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <param name="PivotTableSource">The pivottable source for a pivotchart</param>    
    /// <returns>The chart</returns>
    public ExcelBarChart AddBarChart(string Name, eBarChartType ChartType, ExcelPivotTable PivotTableSource)
    {
        return (ExcelBarChart)this.AddAllChartTypes(Name, (eChartType)ChartType, PivotTableSource);
    }

    /// <summary>
    /// Adds a new pie chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <returns>The chart</returns>    
    public ExcelPieChart AddPieChart(string Name, ePieChartType ChartType)
    {
        return (ExcelPieChart)this.AddAllChartTypes(Name, (eChartType)ChartType, null);
    }

    /// <summary>
    /// Adds a new pie chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <param name="PivotTableSource">The pivottable source for a pivotchart</param>    
    /// <returns>The chart</returns>
    public ExcelPieChart AddPieChart(string Name, ePieChartType ChartType, ExcelPivotTable PivotTableSource)
    {
        return (ExcelPieChart)this.AddAllChartTypes(Name, (eChartType)ChartType, PivotTableSource);
    }

    /// <summary>
    /// Adds a new doughnut chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <param name="PivotTableSource">The pivottable source for a pivotchart</param>    
    /// <returns>The chart</returns>
    public ExcelDoughnutChart AddDoughnutChart(string Name, eDoughnutChartType ChartType, ExcelPivotTable PivotTableSource)
    {
        return (ExcelDoughnutChart)this.AddAllChartTypes(Name, (eChartType)ChartType, PivotTableSource);
    }

    /// <summary>
    /// Adds a new doughnut chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <returns>The chart</returns>    
    public ExcelDoughnutChart AddDoughnutChart(string Name, eDoughnutChartType ChartType)
    {
        return (ExcelDoughnutChart)this.AddAllChartTypes(Name, (eChartType)ChartType, null);
    }

    /// <summary>
    /// Adds a new line chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <returns>The chart</returns>    
    public ExcelOfPieChart AddOfPieChart(string Name, eOfPieChartType ChartType)
    {
        return (ExcelOfPieChart)this.AddAllChartTypes(Name, (eChartType)ChartType, null);
    }

    /// <summary>
    /// Add a new pie of pie or bar of pie chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <param name="PivotTableSource">The pivottable source for a pivotchart</param>    
    /// <returns>The chart</returns>
    public ExcelOfPieChart AddOfPieChart(string Name, eOfPieChartType ChartType, ExcelPivotTable PivotTableSource)
    {
        return (ExcelOfPieChart)this.AddAllChartTypes(Name, (eChartType)ChartType, PivotTableSource);
    }

    /// <summary>
    /// Adds a new bubble chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <returns>The chart</returns>    
    public ExcelBubbleChart AddBubbleChart(string Name, eBubbleChartType ChartType)
    {
        return (ExcelBubbleChart)this.AddAllChartTypes(Name, (eChartType)ChartType, null);
    }

    /// <summary>
    /// Adds a new bubble chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <param name="PivotTableSource">The pivottable source for a pivotchart</param>    
    /// <returns>The chart</returns>
    public ExcelBubbleChart AddBubbleChart(string Name, eBubbleChartType ChartType, ExcelPivotTable PivotTableSource)
    {
        return (ExcelBubbleChart)this.AddAllChartTypes(Name, (eChartType)ChartType, PivotTableSource);
    }

    /// <summary>
    /// Adds a new scatter chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <param name="PivotTableSource">The pivottable source for a pivotchart</param>    
    /// <returns>The chart</returns>
    public ExcelScatterChart AddScatterChart(string Name, eScatterChartType ChartType, ExcelPivotTable PivotTableSource)
    {
        return (ExcelScatterChart)this.AddAllChartTypes(Name, (eChartType)ChartType, PivotTableSource);
    }

    /// <summary>
    /// Adds a new scatter chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <returns>The chart</returns>    
    public ExcelScatterChart AddScatterChart(string Name, eScatterChartType ChartType)
    {
        return (ExcelScatterChart)this.AddAllChartTypes(Name, (eChartType)ChartType, null);
    }

    /// <summary>
    /// Adds a new radar chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <param name="PivotTableSource">The pivottable source for a pivotchart</param>    
    /// <returns>The chart</returns>
    public ExcelRadarChart AddRadarChart(string Name, eRadarChartType ChartType, ExcelPivotTable PivotTableSource)
    {
        return (ExcelRadarChart)this.AddAllChartTypes(Name, (eChartType)ChartType, PivotTableSource);
    }

    /// <summary>
    /// Adds a new radar chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <returns>The chart</returns>    
    public ExcelRadarChart AddRadarChart(string Name, eRadarChartType ChartType)
    {
        return (ExcelRadarChart)this.AddAllChartTypes(Name, (eChartType)ChartType, null);
    }

    /// <summary>
    /// Adds a new surface chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <param name="PivotTableSource">The pivottable source for a pivotchart</param>    
    /// <returns>The chart</returns>
    public ExcelSurfaceChart AddSurfaceChart(string Name, eSurfaceChartType ChartType, ExcelPivotTable PivotTableSource)
    {
        return (ExcelSurfaceChart)this.AddAllChartTypes(Name, (eChartType)ChartType, PivotTableSource);
    }

    /// <summary>
    /// Adds a new surface chart to the worksheet.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ChartType">Type of chart</param>
    /// <returns>The chart</returns>    
    public ExcelSurfaceChart AddSurfaceChart(string Name, eSurfaceChartType ChartType)
    {
        return (ExcelSurfaceChart)this.AddAllChartTypes(Name, (eChartType)ChartType, null);
    }

    /// <summary>
    /// Adds a picture to the worksheet
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ImageFile">The image file</param>
    /// <returns>A picture object</returns>
    public ExcelPicture AddPicture(string Name, FileInfo ImageFile)
    {
        return this.AddPicture(Name, ImageFile, null);
    }

    /// <summary>
    /// Adds a picture to the worksheet
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ImageFile">The image file</param>
    /// <param name="Hyperlink">Picture Hyperlink</param>
    /// <returns>A picture object</returns>
    public ExcelPicture AddPicture(string Name, FileInfo ImageFile, Uri Hyperlink)
    {
        this.ValidatePictureFile(Name, ImageFile);
        XmlElement drawNode = this.CreateDrawingXml(eEditAs.OneCell);
        ePictureType type = PictureStore.GetPictureType(ImageFile.Extension);
        ExcelPicture? pic = new ExcelPicture(this, drawNode, Hyperlink, type);
        pic.LoadImage(new FileStream(ImageFile.FullName, FileMode.Open, FileAccess.Read), type);
        this.AddPicture(Name, pic);

        return pic;
    }

    /// <summary>
    /// Adds a picture to the worksheet using a stream. EPPlus will identify the type of image automatically.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="PictureStream">An stream image.</param>
    /// <returns>A picture object</returns>
    public ExcelPicture AddPicture(string Name, Stream PictureStream)
    {
        return this.AddImageInternal(Name, PictureStream, null, null);
    }

    /// <summary>
    /// Adds a picture to the worksheet from a stream. EPPlus will identify the type of image automatically.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="PictureStream">An stream image.</param>        
    /// <param name="Hyperlink">The Picture Hyperlink</param>
    /// <returns>A picture object</returns>
    public ExcelPicture AddPicture(string Name, Stream PictureStream, Uri Hyperlink)
    {
        return this.AddImageInternal(Name, PictureStream, null, Hyperlink);
    }

    /// <summary>
    /// Adds a picture to the worksheet
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="PictureStream">An stream image.</param>
    /// <param name="PictureType">The type of image.  A null value means that EPPlus will identify the type of image automatically.</param>
    /// <returns>A picture object</returns>
    [Obsolete("This overload is deprecated, please use AddPicture(string, Stream) instead.")]
    public ExcelPicture AddPicture(string Name, Stream PictureStream, ePictureType? PictureType)
    {
        return this.AddPicture(Name, PictureStream, PictureType, null);
    }

    /// <summary>
    /// Adds a picture to the worksheet
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="pictureStream">An stream image.</param>
    /// <param name="pictureType">The type of image. A null value means that EPPlus will identify the type of image automatically.</param>
    /// <param name="Hyperlink">Picture Hyperlink</param>
    /// <returns>A picture object</returns>
    [Obsolete("This overload is deprecated, please use AddPicture(string, Stream, Uri) instead.")]
    public ExcelPicture AddPicture(string Name, Stream pictureStream, ePictureType? pictureType, Uri Hyperlink)
    {
        return this.AddImageInternal(Name, pictureStream, pictureType, Hyperlink);
    }

    private ExcelPicture AddImageInternal(string Name, Stream pictureStream, ePictureType? pictureType, Uri Hyperlink)
    {
        if (pictureStream == null)
        {
            throw new ArgumentNullException("Stream cannot be null");
        }

        if (!pictureStream.CanRead || !pictureStream.CanSeek)
        {
            throw new IOException("Stream must be readable and seekable");
        }

        pictureType ??= ImageReader.GetPictureType(pictureStream, true);

        XmlElement drawNode = this.CreateDrawingXml(eEditAs.OneCell);
        ExcelPicture? pic = new ExcelPicture(this, drawNode, Hyperlink, pictureType.Value);
        pic.LoadImage(pictureStream, pictureType.Value);
        this.AddPicture(Name, pic);

        return pic;
    }

    internal ExcelGroupShape AddGroupDrawing()
    {
        XmlElement drawNode = this.CreateDrawingXml(eEditAs.OneCell);
        ExcelGroupShape? grp = new ExcelGroupShape(this, drawNode);
        grp.Name = $"Group {grp.Id}";
        this._drawingsList.Add(grp);
        this._drawingNames.Add(grp.Name, this._drawingsList.Count - 1);

        return grp;
    }

    #region AddPictureAsync

#if !NET35 && !NET40
    /// <summary>
    /// Adds a picture to the worksheet
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ImageFile">The image file</param>
    /// <returns>A picture object</returns>
    public async Task<ExcelPicture> AddPictureAsync(string Name, FileInfo ImageFile)
    {
        return await this.AddPictureAsync(Name, ImageFile, null);
    }

    /// <summary>
    /// Adds a picture to the worksheet
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ImageFile">The image file</param>
    /// <param name="Hyperlink">Picture Hyperlink</param>
    /// <returns>A picture object</returns>
    public async Task<ExcelPicture> AddPictureAsync(string Name, FileInfo ImageFile, Uri Hyperlink)
    {
        this.ValidatePictureFile(Name, ImageFile);
        XmlElement drawNode = this.CreateDrawingXml(eEditAs.OneCell);
        ePictureType type = PictureStore.GetPictureType(ImageFile.Extension);
        ExcelPicture? pic = new ExcelPicture(this, drawNode, Hyperlink, type);
        await pic.LoadImageAsync(new FileStream(ImageFile.FullName, FileMode.Open, FileAccess.Read), type);
        this.AddPicture(Name, pic);

        return pic;
    }

    /// <summary>
    /// Adds a picture to the worksheet
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ImagePath">The path to the image file</param>
    /// <returns>A picture object</returns>
    public async Task<ExcelPicture> AddPictureAsync(string Name, string ImagePath)
    {
        return await this.AddPictureAsync(Name, new FileInfo(ImagePath), null);
    }

    /// <summary>
    /// Adds a picture to the worksheet
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="ImagePath">The path to the image file</param>
    /// <param name="Hyperlink">Picture Hyperlink</param>
    /// <returns>A picture object</returns>
    public async Task<ExcelPicture> AddPictureAsync(string Name, string ImagePath, Uri Hyperlink)
    {
        return await this.AddPictureAsync(Name, new FileInfo(ImagePath), Hyperlink);
    }

    /// <summary>
    /// Adds a picture to the worksheet from a stream. EPPlus will identify the type of image automatically.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="PictureStream">An stream image.</param>        
    /// <returns>A picture object</returns>
    public async Task<ExcelPicture> AddPictureAsync(string Name, Stream PictureStream)
    {
        return await this.AddPictureInternalAsync(Name, PictureStream, null, null);
    }

    /// <summary>
    /// Adds a picture to the worksheet from a stream. EPPlus will identify the type of image automatically.
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="PictureStream">An stream image.</param>        
    /// <param name="Hyperlink">The Picture Hyperlink</param>
    /// <returns>A picture object</returns>
    public async Task<ExcelPicture> AddPictureAsync(string Name, Stream PictureStream, Uri Hyperlink)
    {
        return await this.AddPictureInternalAsync(Name, PictureStream, null, Hyperlink);
    }

    /// <summary>
    /// Adds a picture to the worksheet
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="PictureStream">An stream image.</param>
    /// <param name="PictureType">The type of image. A null value means that EPPlus will identify the type of image automatically.</param>
    /// <returns>A picture object</returns>
    [Obsolete("This overload is deprecated, please use AddPictureAsync(string, Stream) instead.")]
    public async Task<ExcelPicture> AddPictureAsync(string Name, Stream PictureStream, ePictureType? PictureType)
    {
        return await this.AddPictureInternalAsync(Name, PictureStream, PictureType, null);
    }

    /// <summary>
    /// Adds a picture to the worksheet
    /// </summary>
    /// <param name="Name"></param>
    /// <param name="pictureStream">An stream image.</param>
    /// <param name="pictureType">The type of image. A null value means that EPPlus will identify the type of image automatically.</param>
    /// <param name="Hyperlink">The Picture Hyperlink</param>
    /// <returns>A picture object</returns>
    [Obsolete("This overload is deprecated, please use AddPictureAsync(string, Stream, Uri) instead.")]
    public async Task<ExcelPicture> AddPictureAsync(string Name, Stream pictureStream, ePictureType? pictureType, Uri Hyperlink)
    {
        return await this.AddPictureInternalAsync(Name, pictureStream, pictureType, Hyperlink);
    }

    private async Task<ExcelPicture> AddPictureInternalAsync(string Name, Stream pictureStream, ePictureType? pictureType, Uri Hyperlink)
    {
        if (pictureStream == null)
        {
            throw new ArgumentNullException("Stream cannot be null");
        }

        if (!pictureStream.CanRead || !pictureStream.CanSeek)
        {
            throw new IOException("Stream must be readable and seekable");
        }

        XmlElement drawNode = this.CreateDrawingXml(eEditAs.OneCell);
        pictureType ??= await ImageReader.GetPictureTypeAsync(pictureStream);

        ExcelPicture? pic = new ExcelPicture(this, drawNode, Hyperlink, pictureType.Value);
        await pic.LoadImageAsync(pictureStream, pictureType.Value);
        this.AddPicture(Name, pic);

        return pic;
    }
#endif

    #endregion

    private void AddPicture(string Name, ExcelPicture pic)
    {
        pic.Name = Name;
        this._drawingsList.Add(pic);
        this._drawingNames.Add(Name, this._drawingsList.Count - 1);
    }

    /// <summary>
    /// Adds a picture to the worksheet
    /// </summary>
    /// <param name="Name">The name of the drawing object</param>
    /// <param name="ImagePath">The path to the image file</param>
    /// <returns>A picture object</returns>
    public ExcelPicture AddPicture(string Name, string ImagePath)
    {
        if (string.IsNullOrEmpty(ImagePath) == false)
        {
            return this.AddPicture(Name, new FileInfo(ImagePath), null);
        }

        throw new Exception("AddPicture: Image path can't be null");
    }

    /// <summary>
    /// Adds a picture to the worksheet
    /// </summary>
    /// <param name="Name">The name of the drawing object</param>
    /// <param name="ImagePath">The path to the image file</param>
    /// <param name="Hyperlink">Picture Hyperlink</param>
    /// <returns>A picture object</returns>
    public ExcelPicture AddPicture(string Name, string ImagePath, ExcelHyperLink Hyperlink)
    {
        if (string.IsNullOrEmpty(ImagePath) == false)
        {
            return this.AddPicture(Name, new FileInfo(ImagePath), Hyperlink);
        }

        throw new Exception("AddPicture: Image path can't be null");
    }

    private void ValidatePictureFile(string Name, FileInfo ImageFile)
    {
        if (this.Worksheet is ExcelChartsheet && this._drawingsList.Count > 0)
        {
            throw new InvalidOperationException("Chart worksheets can't have more than one drawing");
        }

        if (ImageFile == null)
        {
            throw new Exception("AddPicture: ImageFile can't be null");
        }

        if (!ImageFile.Exists)
        {
            throw new FileNotFoundException("Cant find file.", ImageFile.FullName);
        }

        if (this._drawingNames.ContainsKey(Name))
        {
            throw new Exception("Name already exists in the drawings collection");
        }
    }

    /// <summary>
    /// Adds a new chart using an crtx template
    /// </summary>
    /// <param name="crtxFile">The crtx file</param>
    /// <param name="name">The name of the chart</param>
    /// <returns>The new chart</returns>
    public ExcelChart AddChartFromTemplate(FileInfo crtxFile, string name)
    {
        return this.AddChartFromTemplate(crtxFile, name, null);
    }

    /// <summary>
    /// Adds a new chart using an crtx template
    /// </summary>
    /// <param name="crtxFile">The crtx file</param>
    /// <param name="name">The name of the chart</param>
    /// <param name="pivotTableSource">Pivot table source, if the chart is a pivottable</param>
    /// <returns>The new chart</returns>
    public ExcelChart AddChartFromTemplate(FileInfo crtxFile, string name, ExcelPivotTable pivotTableSource)
    {
        if (!crtxFile.Exists)
        {
            throw new FileNotFoundException($"{crtxFile.FullName} cannot be found.");
        }

        FileStream fs = null;

        try
        {
            fs = crtxFile.Open(FileMode.Open, FileAccess.Read, FileShare.Read);

            return this.AddChartFromTemplate(fs, name);
        }
        catch
        {
            throw;
        }
        finally
        {
            if (fs != null)
            {
                fs.Close();
            }
        }
    }

    /// <summary>
    /// Adds a new chart using an crtx template
    /// </summary>
    /// <param name="crtxStream">The crtx file as a stream</param>
    /// <param name="name">The name of the chart</param>
    /// <returns>The new chart</returns>
    public ExcelChart AddChartFromTemplate(Stream crtxStream, string name)
    {
        return this.AddChartFromTemplate(crtxStream, name, null);
    }

    /// <summary>
    /// Adds a new chart using an crtx template
    /// </summary>
    /// <param name="crtxStream">The crtx file as a stream</param>
    /// <param name="name">The name of the chart</param>
    /// <param name="pivotTableSource">Pivot table source, if the chart is a pivottable</param>
    /// <returns>The new chart</returns>
    public ExcelChart AddChartFromTemplate(Stream crtxStream, string name, ExcelPivotTable pivotTableSource)
    {
        if (this.Worksheet is ExcelChartsheet && this._drawingsList.Count > 0)
        {
            throw new InvalidOperationException("Chart worksheets can't have more than one drawing");
        }

        CrtxTemplateHelper.LoadCrtx(crtxStream,
                                    out XmlDocument chartXml,
                                    out XmlDocument styleXml,
                                    out XmlDocument colorsXml,
                                    out ZipPackagePart themePart,
                                    "The crtx stream");

        if (chartXml == null)
        {
            throw new InvalidDataException("Crtx file is corrupt.");
        }

        XmlHelper? chartXmlHelper = XmlHelperFactory.Create(this.NameSpaceManager, chartXml.DocumentElement);
        XmlNode? serNode = chartXmlHelper.GetNode("/c:chartSpace/c:chart/c:plotArea/*[substring(name(), string-length(name()) - 4) = 'Chart']/c:ser");

        if (serNode != null)
        {
            this._seriesTemplateXml = serNode.InnerXml;
            _ = serNode.ParentNode.RemoveChild(serNode);
        }

        XmlElement drawNode = this.CreateDrawingXml(eEditAs.TwoCell);
        eChartType? chartType = ExcelChart.GetChartTypeFromNodeName(GetChartNodeName(chartXmlHelper));
        ExcelChart? chart = ExcelChart.GetNewChart(this, drawNode, chartType, null, pivotTableSource, chartXml);

        chart.Name = name;
        this._drawingsList.Add(chart);
        this._drawingNames.Add(name, this._drawingsList.Count - 1);
        eChartStyle chartStyle = chart.Style;

        if (chartStyle == eChartStyle.None)
        {
            chartStyle = eChartStyle.Style2;
        }

        if (themePart != null)
        {
            chart.StyleManager.LoadThemeOverrideXml(themePart);
        }

        _ = chart.StyleManager.LoadStyleXml(styleXml, chartStyle, colorsXml);

        return chart;
    }

    private static string GetChartNodeName(XmlHelper xmlHelper)
    {
        XmlNode? ploterareaNode = xmlHelper.GetNode(ExcelChart.plotAreaPath);

        foreach (XmlNode node in ploterareaNode?.ChildNodes)
        {
            if (node.LocalName.EndsWith("Chart"))
            {
                return node.LocalName;
            }
        }

        return "";
    }

    /// <summary>
    /// Adds a new shape to the worksheet
    /// </summary>
    /// <param name="Name">Name</param>
    /// <param name="Style">Shape style</param>
    /// <returns>The shape object</returns>
    public ExcelShape AddShape(string Name, eShapeStyle Style)
    {
        if (this.Worksheet is ExcelChartsheet && this._drawingsList.Count > 0)
        {
            throw new InvalidOperationException("Chart worksheets can't have more than one drawing");
        }

        if (this._drawingNames.ContainsKey(Name))
        {
            throw new Exception("Name already exists in the drawings collection");
        }

        XmlElement drawNode = this.CreateDrawingXml();

        ExcelShape shape = new ExcelShape(this, drawNode, Style);
        shape.Name = Name;
        this._drawingsList.Add(shape);
        this._drawingNames.Add(Name, this._drawingsList.Count - 1);

        return shape;
    }

    #region Add Slicers

    /// <summary>
    /// Adds a slicer to a table column
    /// </summary>
    /// <param name="TableColumn">The table column</param>
    /// <returns>The slicer drawing</returns>
    public ExcelTableSlicer AddTableSlicer(ExcelTableColumn TableColumn)
    {
        if (this.Worksheet is ExcelChartsheet && this._drawingsList.Count > 0)
        {
            throw new InvalidOperationException("Chart worksheets can't have more than one drawing");
        }

        if (TableColumn.Table.AutoFilter.Columns[TableColumn.Position] == null)
        {
            _ = TableColumn.Table.AutoFilter.Columns.AddValueFilterColumn(TableColumn.Position);
        }

        XmlElement drawNode = this.CreateDrawingXml();
        ExcelTableSlicer? slicer = new ExcelTableSlicer(this, drawNode, TableColumn) { EditAs = eEditAs.OneCell, };
        slicer.SetSize(192, 260);

        this._drawingsList.Add(slicer);
        this._drawingNames.Add(slicer.Name, this._drawingsList.Count - 1);

        return slicer;
    }

    /// <summary>
    /// Adds a slicer to a pivot table field
    /// </summary>
    /// <param name="Field">The pivot table field</param>
    /// <returns>The slicer drawing</returns>
    internal ExcelPivotTableSlicer AddPivotTableSlicer(ExcelPivotTableField Field)
    {
        if (this.Worksheet is ExcelChartsheet && this._drawingsList.Count > 0)
        {
            throw new InvalidOperationException("Chart worksheets can't have more than one drawing");
        }

        if (!string.IsNullOrEmpty(Field.Cache.Formula))
        {
            throw new InvalidOperationException("Can't add a slicer to a calculated field");
        }

        if (Field._pivotTable.CacheId == 0)
        {
            _ = Field._pivotTable.ChangeCacheId(0); //Slicers can for some reason not have a cache id of 0.
        }

        XmlElement drawNode = this.CreateDrawingXml();
        ExcelPivotTableSlicer? slicer = new ExcelPivotTableSlicer(this, drawNode, Field) { EditAs = eEditAs.OneCell, };
        slicer.SetSize(192, 260);
        this._drawingsList.Add(slicer);
        this._drawingNames.Add(slicer.Name, this._drawingsList.Count - 1);

        return slicer;
    }

    #endregion

    ///// <summary>
    ///// Adds a line connectin two shapes
    ///// </summary>
    ///// <param name="Name">The Name</param>
    ///// <param name="Style">The connectorStyle</param>
    ///// <param name="StartShape">The starting shape to connect</param>
    ///// <param name="EndShape">The ending shape to connect</param>
    ///// <returns></returns>
    //public ExcelConnectionShape AddShape(string Name, eShapeConnectorStyle Style, ExcelShape StartShape, ExcelShape EndShape)
    //{
    //    if (Worksheet is ExcelChartsheet && _drawings.Count > 0)
    //    {
    //        throw new InvalidOperationException("Chart worksheets can't have more than one drawing");
    //    }
    //    if (_drawingNames.ContainsKey(Name))
    //    {
    //        throw new Exception("Name already exists in the drawings collection");
    //    }
    //    var drawNode = CreateDrawingXml();

    //    var shape = new ExcelConnectionShape(this, drawNode, Style, StartShape, EndShape);

    //    shape.Name = Name;
    //    _drawings.Add(shape);
    //    _drawingNames.Add(Name, _drawings.Count - 1);
    //    return shape;
    //}

    /// <summary>
    /// Adds a new shape to the worksheet
    /// </summary>
    /// <param name="Name">Name</param>
    /// <param name="Source">Source shape</param>
    /// <returns>The shape object</returns>
    public ExcelShape AddShape(string Name, ExcelShape Source)
    {
        if (this.Worksheet is ExcelChartsheet && this._drawingsList.Count > 0)
        {
            throw new InvalidOperationException("Chart worksheets can't have more than one drawing");
        }

        if (this._drawingNames.ContainsKey(Name))
        {
            throw new Exception("Name already exists in the drawings collection");
        }

        XmlElement drawNode = this.CreateDrawingXml();
        drawNode.InnerXml = Source.TopNode.InnerXml;

        ExcelShape shape = new ExcelShape(this, drawNode);
        shape.Name = Name;
        shape.Style = Source.Style;
        this._drawingsList.Add(shape);
        this._drawingNames.Add(Name, this._drawingsList.Count - 1);

        return shape;
    }

    #region Form Controls

    /// <summary>
    /// Adds a form control to the worksheet
    /// </summary>
    /// <param name="Name">The name</param>
    /// <param name="ControlType">The type of control</param>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException">Chart sheets cannot have controls</exception>
    /// <exception cref="ArgumentException">Drawing names must be unique</exception>
    public ExcelControl AddControl(string Name, eControlType ControlType)
    {
        if (this.Worksheet is ExcelChartsheet && this._drawingsList.Count > 0)
        {
            throw new InvalidOperationException("Chart worksheets can't have more than one drawing");
        }

        if (this._drawingNames.ContainsKey(Name))
        {
            throw new ArgumentException("Name already exists in the drawings collection");
        }

        XmlElement drawNode = this.CreateDrawingXml(eEditAs.TwoCell, true);

        ExcelControl control = ControlFactory.CreateControl(ControlType, this, drawNode, Name);
        control.EditAs = ExcelControl.GetControlEditAs(ControlType);
        this._drawingsList.Add(control);
        this._drawingNames.Add(Name, this._drawingsList.Count - 1);

        return control;
    }

    /// <summary>
    /// Adds a button form control to the worksheet
    /// </summary>
    /// <param name="Name">The name of the button</param>
    /// <returns>The button form control</returns>
    public ExcelControlButton AddButtonControl(string Name)
    {
        return (ExcelControlButton)this.AddControl(Name, eControlType.Button);
    }

    /// <summary>
    /// Adds a checkbox form control to the worksheet
    /// </summary>
    /// <param name="Name">The name of the checkbox control</param>
    /// <returns>The checkbox form control</returns>
    public ExcelControlCheckBox AddCheckBoxControl(string Name)
    {
        return (ExcelControlCheckBox)this.AddControl(Name, eControlType.CheckBox);
    }

    /// <summary>
    /// Adds a radio button form control to the worksheet
    /// </summary>
    /// <param name="Name">The name of the radio button control</param>
    /// <returns>The radio button form control</returns>
    public ExcelControlRadioButton AddRadioButtonControl(string Name)
    {
        return (ExcelControlRadioButton)this.AddControl(Name, eControlType.RadioButton);
    }

    /// <summary>
    /// Adds a list box form control to the worksheet
    /// </summary>
    /// <param name="Name">The name of the list box control</param>
    /// <returns>The list box form control</returns>
    public ExcelControlListBox AddListBoxControl(string Name)
    {
        return (ExcelControlListBox)this.AddControl(Name, eControlType.ListBox);
    }

    /// <summary>
    /// Adds a drop-down form control to the worksheet
    /// </summary>
    /// <param name="Name">The name of the drop-down control</param>
    /// <returns>The drop-down form control</returns>
    public ExcelControlDropDown AddDropDownControl(string Name)
    {
        return (ExcelControlDropDown)this.AddControl(Name, eControlType.DropDown);
    }

    /// <summary>
    /// Adds a group box form control to the worksheet
    /// </summary>
    /// <param name="Name">The name of the group box control</param>
    /// <returns>The group box form control</returns>
    public ExcelControlGroupBox AddGroupBoxControl(string Name)
    {
        return (ExcelControlGroupBox)this.AddControl(Name, eControlType.GroupBox);
    }

    /// <summary>
    /// Adds a label form control to the worksheet
    /// </summary>
    /// <param name="Name">The name of the label control</param>
    /// <returns>The label form control</returns>
    public ExcelControlLabel AddLabelControl(string Name)
    {
        return (ExcelControlLabel)this.AddControl(Name, eControlType.Label);
    }

    /// <summary>
    /// Adds a spin button control to the worksheet
    /// </summary>
    /// <param name="Name">The name of the spin button control</param>
    /// <returns>The spin button form control</returns>
    public ExcelControlSpinButton AddSpinButtonControl(string Name)
    {
        return (ExcelControlSpinButton)this.AddControl(Name, eControlType.SpinButton);
    }

    /// <summary>
    /// Adds a scroll bar control to the worksheet
    /// </summary>
    /// <param name="Name">The name of the scroll bar control</param>
    /// <returns>The scroll bar form control</returns>
    public ExcelControlScrollBar AddScrollBarControl(string Name)
    {
        return (ExcelControlScrollBar)this.AddControl(Name, eControlType.ScrollBar);
    }

    #endregion

    private XmlElement CreateDrawingXml(eEditAs topNodeType = eEditAs.TwoCell, bool asAlterniveContent = false)
    {
        if (this.DrawingXml.DocumentElement == null)
        {
            this.DrawingXml.LoadXml(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><xdr:wsDr xmlns:xdr=\"{0}\" xmlns:a=\"{1}\" />",
                                                  ExcelPackage.schemaSheetDrawings,
                                                  ExcelPackage.schemaDrawings));

            ZipPackage package = this.Worksheet._package.ZipPackage;

            //Check for existing part, issue #100
            int id = this.Worksheet.SheetId;

            do
            {
                this._uriDrawing = new Uri(string.Format("/xl/drawings/drawing{0}.xml", id++), UriKind.Relative);
            } while (package.PartExists(this._uriDrawing));

            this._part = package.CreatePart(this._uriDrawing, "application/vnd.openxmlformats-officedocument.drawing+xml", this._package.Compression);

            StreamWriter streamChart = new StreamWriter(this._part.GetStream(FileMode.Create, FileAccess.Write));
            this.DrawingXml.Save(streamChart);
            streamChart.Close();
            ZipPackage.Flush();

            this._drawingRelation = this.Worksheet.Part.CreateRelationship(UriHelper.GetRelativeUri(this.Worksheet.WorksheetUri, this._uriDrawing),
                                                                           TargetMode.Internal,
                                                                           ExcelPackage.schemaRelationships + "/drawing");

            XmlElement e = (XmlElement)this.Worksheet.CreateNode("d:drawing");
            _ = e.SetAttribute("id", ExcelPackage.schemaRelationships, this._drawingRelation.Id);

            ZipPackage.Flush();
        }

        XmlNode colNode = this._drawingsXml.SelectSingleNode("//xdr:wsDr", this.NameSpaceManager);

        string? topElementname = $"{topNodeType.ToEnumString()}Anchor";
        XmlElement drawNode = this._drawingsXml.CreateElement("xdr", topElementname, ExcelPackage.schemaSheetDrawings);

        if (asAlterniveContent)
        {
            XmlElement? acNode = (XmlElement)this._drawingsXml.CreateElement("mc", "AlternateContent", ExcelPackage.schemaMarkupCompatibility);
            acNode.SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);
            acNode.InnerXml = "<mc:Choice Requires=\"a14\" xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\"></mc:Choice><mc:Fallback/>";
            _ = acNode.FirstChild.AppendChild(drawNode);
            _ = colNode.AppendChild(acNode);
        }
        else
        {
            _ = colNode.AppendChild(drawNode);
        }

        if (topNodeType == eEditAs.OneCell || topNodeType == eEditAs.TwoCell)
        {
            //Add from position Element;
            XmlElement fromNode = this._drawingsXml.CreateElement("xdr", "from", ExcelPackage.schemaSheetDrawings);
            _ = drawNode.AppendChild(fromNode);
            fromNode.InnerXml = "<xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff>";
        }
        else
        {
            //Add from position Element;
            XmlElement posNode = this._drawingsXml.CreateElement("xdr", "pos", ExcelPackage.schemaSheetDrawings);
            posNode.SetAttribute("x", "0");
            posNode.SetAttribute("y", "0");
            _ = drawNode.AppendChild(posNode);
        }

        if (topNodeType == eEditAs.TwoCell)
        {
            //Add to position Element;
            XmlElement toNode = this._drawingsXml.CreateElement("xdr", "to", ExcelPackage.schemaSheetDrawings);
            _ = drawNode.AppendChild(toNode);
            toNode.InnerXml = "<xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff>";
        }
        else
        {
            //Add from position Element;
            XmlElement posNode = this._drawingsXml.CreateElement("xdr", "ext", ExcelPackage.schemaSheetDrawings);
            posNode.SetAttribute("cx", "6072876");
            posNode.SetAttribute("cy", "9299263");
            _ = drawNode.AppendChild(posNode);
        }

        return drawNode;
    }

    #endregion

    #region Remove methods

    /// <summary>
    /// Removes a drawing.
    /// </summary>
    /// <param name="Index">The index of the drawing</param>
    public void Remove(int Index)
    {
        if (this.Worksheet is ExcelChartsheet && this._drawingsList.Count > 0)
        {
            throw new InvalidOperationException("Can' remove charts from chart worksheets");
        }

        this.RemoveDrawing(Index);
    }

    internal void RemoveDrawing(int Index, bool DeleteXmlNode = true)
    {
        ExcelDrawing? draw = this._drawingsList[Index];

        if (DeleteXmlNode)
        {
            draw.DeleteMe();
        }

        _ = this._drawingNames.Remove(draw.Name);
        _ = this._drawingsList.Remove(draw);
        this.ReIndexNames(Index, -1);
    }

    internal void ReIndexNames(int Index, int increase)
    {
        for (int i = this._drawingsList.Count - 1; i >= Index; i--)
        {
            this._drawingNames[this._drawingsList[i].Name] = i;
        }
    }

    /// <summary>
    /// Removes a drawing.
    /// </summary>
    /// <param name="Drawing">The drawing</param>
    public void Remove(ExcelDrawing Drawing)
    {
        this.Remove(this._drawingNames[Drawing.Name]);
    }

    /// <summary>
    /// Removes a drawing.
    /// </summary>
    /// <param name="Name">The name of the drawing</param>
    public void Remove(string Name)
    {
        this.Remove(this._drawingNames[Name]);
    }

    /// <summary>
    /// Removes all drawings from the collection
    /// </summary>
    public void Clear()
    {
        if (this.Worksheet is ExcelChartsheet && this._drawingsList.Count > 0)
        {
            throw new InvalidOperationException("Can' remove charts from chart worksheets");
        }

        this.ClearDrawings();
    }

    internal void ClearDrawings()
    {
        while (this.Count > 0)
        {
            this.RemoveDrawing(0);
        }
    }

    #endregion

    #region BringToFront & SendToBack

    internal void BringToFront(ExcelDrawing drawing)
    {
        int index = this._drawingsList.IndexOf(drawing);
        int endIndex = this._drawingsList.Count - 1;

        if (index == endIndex)
        {
            return;
        }

        //Move in Xml
        XmlNode? parentNode = drawing.TopNode.ParentNode;
        _ = parentNode.RemoveChild(drawing.TopNode);
        _ = parentNode.InsertAfter(drawing.TopNode, parentNode.LastChild);

        //Move in list 
        this._drawingsList.RemoveAt(index);
        this._drawingsList.Insert(endIndex, drawing);

        //Reindex dictionary
        this._drawingNames[drawing.Name] = endIndex;

        for (int i = index + 0; i < endIndex; i++)
        {
            this._drawingNames[this._drawingsList[i].Name]--;
        }
    }

    internal void SendToBack(ExcelDrawing drawing)
    {
        int index = this._drawingsList.IndexOf(drawing);

        if (index == 0)
        {
            return;
        }

        //Move in Xml
        XmlNode? parentNode = drawing.TopNode.ParentNode;
        _ = parentNode.RemoveChild(drawing.TopNode);
        _ = parentNode.InsertBefore(drawing.TopNode, parentNode.FirstChild);

        //Move in list 
        this._drawingsList.RemoveAt(index);
        this._drawingsList.Insert(0, drawing);

        //Reindex dictionary
        this._drawingNames[drawing.Name] = 0;

        for (int i = 1; i <= index; i++)
        {
            this._drawingNames[this._drawingsList[i].Name]++;
        }
    }

    #endregion

    internal void AdjustWidth(double[,] pos)
    {
        int ix = 0;

        //Now set the size for all drawings depending on the editAs property.
        foreach (ExcelDrawing d in this)
        {
            if (d.EditAs != eEditAs.TwoCell)
            {
                if (d.EditAs == eEditAs.Absolute)
                {
                    d.SetPixelLeft(pos[ix, 0]);
                }

                d.SetPixelWidth(pos[ix, 1]);
            }

            if (d is ExcelGroupShape gr)
            {
                gr.AdjustChildrenForResizeColumn(pos[ix, 0]);
            }

            ix++;
        }
    }

    internal void AdjustHeight(double[,] pos)
    {
        int ix = 0;

        //Now set the size for all drawings depending on the editAs property.
        foreach (ExcelDrawing d in this)
        {
            if (d.EditAs != eEditAs.TwoCell)
            {
                if (d.EditAs == eEditAs.Absolute)
                {
                    d.SetPixelTop(pos[ix, 0]);
                }

                d.SetPixelHeight(pos[ix, 1]);
            }

            if (d is ExcelGroupShape gr)
            {
                gr.AdjustChildrenForResizeRow(pos[ix, 0]);
            }

            ix++;
        }
    }

    internal double[,] GetDrawingWidths()
    {
        double[,] pos = new double[this.Count, 2];
        int ix = 0;

        //Save the size for all drawings
        foreach (ExcelDrawing d in this)
        {
            pos[ix, 0] = d.GetPixelLeft();
            pos[ix++, 1] = d.GetPixelWidth();
        }

        return pos;
    }

    internal double[,] GetDrawingHeight()
    {
        double[,] pos = new double[this.Count, 2];
        int ix = 0;

        //Save the size for all drawings
        foreach (ExcelDrawing d in this)
        {
            pos[ix, 0] = d.GetPixelTop();
            pos[ix++, 1] = d.GetPixelHeight();
        }

        return pos;
    }

    /// <summary>
    /// Disposes the object
    /// </summary>
    public void Dispose()
    {
        this._drawingsXml = null;
        this._part = null;
        this._drawingNames.Clear();
        this._drawingNames = null;
        this._drawingRelation = null;

        foreach (ExcelDrawing? d in this._drawingsList)
        {
            d.Dispose();
        }

        this._drawingsList.Clear();
        this._drawingsList = null;
    }

    internal ExcelDrawing GetById(int id)
    {
        foreach (ExcelDrawing? d in this._drawingsList)
        {
            if (d.Id == id)
            {
                return d;
            }
        }

        return null;
    }
}