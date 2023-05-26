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
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// A charts plot area
/// </summary>
public class ExcelChartPlotArea :  XmlHelper, IDrawingStyleBase
{
    ExcelChart _firstChart;
    ExcelChart _topChart;
    string _nsPrefix;
    internal ExcelChartPlotArea(XmlNamespaceManager ns, XmlNode node, ExcelChart firstChart, string nsPrefix, ExcelChart topChart=null)
        : base(ns,node)
    {
        this._nsPrefix = nsPrefix;
        if(firstChart._isChartEx)
        {
            this.AddSchemaNodeOrder(new string[] { "plotAreaRegion", "plotSurface", "series", "axis","spPr" },
                                    ExcelDrawing._schemaNodeOrderSpPr);
        }
        else
        {
            this.AddSchemaNodeOrder(new string[] { "areaChart", "area3DChart", "lineChart", "line3DChart", "stockChart", "radarChart", "scatterChart", "pieChart", "pie3DChart", "doughnutChart", "barChart", "bar3DChart", "ofPieChart", "surfaceChart", "surface3DChart", "valAx", "catAx", "dateAx", "serAx", "dTable", "spPr" },
                                    ExcelDrawing._schemaNodeOrderSpPr);
        }

        this._firstChart = firstChart;
        this._topChart = topChart ?? firstChart;
        if (this.TopNode.SelectSingleNode("c:dTable", this.NameSpaceManager) != null)
        {
            this.DataTable = new ExcelChartDataTable(firstChart, this.NameSpaceManager, this.TopNode);
        }
    }

    ExcelChartCollection _chartTypes;
    /// <summary>
    /// If a chart contains multiple chart types (e.g lineChart,BarChart), they end up here.
    /// </summary>
    public ExcelChartCollection ChartTypes  
    {
        get
        {
            if (this._chartTypes == null)
            {
                this._chartTypes = new ExcelChartCollection(this._topChart);
                this._chartTypes.Add(this._firstChart);
                if (this._topChart!= this._firstChart)
                {
                    this._chartTypes.Add(this._topChart);
                }
            }
            return this._chartTypes;
        }
    }
    #region Data table
    /// <summary>
    /// Creates a data table in the plotarea
    /// The datatable can also be accessed via the DataTable propery
    /// <see cref="DataTable"/>
    /// </summary>
    public virtual ExcelChartDataTable CreateDataTable()
    {
        if(this.DataTable!=null)
        {
            throw (new InvalidOperationException("Data table already exists"));
        }

        this.DataTable = new ExcelChartDataTable(this._firstChart, this.NameSpaceManager, this.TopNode);
        this._firstChart.ApplyStyleOnPart(this.DataTable, this._firstChart._styleManager?.Style?.DataTable);
        return this.DataTable;
    }
    /// <summary>
    /// Remove the data table if it's created in the plotarea
    /// </summary>
    public virtual void RemoveDataTable()
    {
        this.DeleteAllNode("c:dTable");
        this.DataTable = null;
    }
    /// <summary>
    /// The data table object.
    /// Use the CreateDataTable method to create a datatable if it does not exist.
    /// <see cref="CreateDataTable"/>
    /// <see cref="RemoveDataTable"/>
    /// </summary>
    public ExcelChartDataTable DataTable { get; private set; } = null;
    #endregion
    ExcelDrawingFill _fill = null;
    /// <summary>
    /// Access to fill properties
    /// </summary>
    public ExcelDrawingFill Fill
    {
        get
        {
            return this._fill ??= new ExcelDrawingFill(this._firstChart,
                                                       this.NameSpaceManager,
                                                       this.TopNode,
                                                       $"{this._nsPrefix}:spPr",
                                                       this.SchemaNodeOrder);
        }
    }
    ExcelDrawingBorder _border = null;
    /// <summary>
    /// Access to border properties
    /// </summary>
    public ExcelDrawingBorder Border
    {
        get
        {
            return this._border ??= new ExcelDrawingBorder(this._firstChart,
                                                           this.NameSpaceManager,
                                                           this.TopNode,
                                                           $"{this._nsPrefix}:spPr/a:ln",
                                                           this.SchemaNodeOrder);
        }
    }
    ExcelDrawingEffectStyle _effect = null;
    /// <summary>
    /// Effects
    /// </summary>
    public ExcelDrawingEffectStyle Effect
    {
        get
        {
            return this._effect ??= new ExcelDrawingEffectStyle(this._firstChart,
                                                                this.NameSpaceManager,
                                                                this.TopNode,
                                                                $"{this._nsPrefix}:spPr/a:effectLst",
                                                                this.SchemaNodeOrder);
        }
    }
    ExcelDrawing3D _threeD = null;
    /// <summary>
    /// 3D properties
    /// </summary>
    public ExcelDrawing3D ThreeD
    {
        get
        {
            return this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:spPr", this.SchemaNodeOrder);
        }
    }
    void IDrawingStyleBase.CreatespPr()
    {
        this.CreatespPrNode();
    }
}