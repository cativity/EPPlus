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
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// Enumerates charttypes 
/// </summary>
public class ExcelChartCollection : IEnumerable<ExcelChart>
{
    List<ExcelChart> _list = new List<ExcelChart>();
    ExcelChart _topChart;
    internal ExcelChartCollection(ExcelChart chart)
    {
        this._topChart = chart;
    }
    internal void Add(ExcelChart chart)
    {
        this._list.Add(chart);
    }
    #region Add charts
    /// <summary>
    /// Add a new charttype to the chart
    /// </summary>
    /// <param name="chartType">The type of the new chart</param>
    /// <returns></returns>
    public ExcelChart Add(eChartType chartType)
    {
        if (this._topChart.PivotTableSource != null)
        {
            throw new InvalidOperationException("Cannot add other chart types to a pivot chart");
        }
        else if(this._topChart._isChartEx)
        {
            throw new InvalidOperationException("Extended charts cannot be combined with other chart types");
        }
        else if (ExcelChart.IsType3D(chartType) || this._list[0].IsType3D())
        {
            throw new InvalidOperationException("3D charts cannot be combined with other chart types");
        }

        XmlNode? prependingChartNode = this._list[this._list.Count - 1].TopNode;
        ExcelChart? chart = ExcelChart.GetNewChart(this._topChart.WorkSheet.Drawings, this._topChart.TopNode, chartType, this._topChart, null);

        this._list.Add((ExcelChart)chart);
        return chart;
    }
    /// <summary>
    /// Adds a new line chart to the chart
    /// </summary>
    /// <param name="chartType">The type of the new chart</param>
    /// <returns>The chart</returns>
    public ExcelLineChart AddLineChart(eLineChartType chartType)
    {
        return (ExcelLineChart)this.Add((eChartType)chartType);
    }
    /// <summary>
    /// Adds a new bar chart to the chart
    /// </summary>
    /// <param name="chartType">The type of the new chart</param>
    /// <returns>The chart</returns>
    public ExcelBarChart AddBarChart(eBarChartType chartType)
    {
        return (ExcelBarChart)this.Add((eChartType)chartType);
    }
    /// <summary>
    /// Adds a new area chart to the chart
    /// </summary>
    /// <param name="chartType">The type of the new chart</param>
    /// <returns>The chart</returns>
    public ExcelAreaChart AddAreaChart(eAreaChartType chartType)
    {
        return (ExcelAreaChart)this.Add((eChartType)chartType);
    }
    /// <summary>
    /// Adds a new pie chart to the chart
    /// </summary>
    /// <param name="chartType">The type of the new chart</param>
    /// <returns>The chart</returns>
    public ExcelPieChart AddPieChart(ePieChartType chartType)
    {
        return (ExcelPieChart)this.Add((eChartType)chartType);
    }
    /// <summary>
    /// Adds a new column of pie- or bar of pie chart to the chart
    /// </summary>
    /// <param name="chartType">The type of the new chart</param>
    /// <returns>The chart</returns>
    public ExcelOfPieChart AddOfPieChart(eOfPieChartType chartType)
    {
        return (ExcelOfPieChart)this.Add((eChartType)chartType);
    }
    /// <summary>
    /// Adds a new doughnut chart to the chart
    /// </summary>
    /// <param name="chartType">The type of the new chart</param>
    /// <returns>The chart</returns>
    public ExcelDoughnutChart AddDoughnutChart(eDoughnutChartType chartType)
    {
        return (ExcelDoughnutChart)this.Add((eChartType)chartType);
    }
    /// <summary>
    /// Adds a new radar chart to the chart
    /// </summary>
    /// <param name="chartType">The type of the new chart</param>
    /// <returns>The chart</returns>
    public ExcelRadarChart AddRadarChart(eRadarChartType chartType)
    {
        return (ExcelRadarChart)this.Add((eChartType)chartType);
    }
    /// <summary>
    /// Adds a new scatter chart to the chart
    /// </summary>
    /// <param name="chartType">The type of the new chart</param>
    /// <returns>The chart</returns>
    public ExcelScatterChart AddScatterChart(eScatterChartType chartType)
    {
        return (ExcelScatterChart)this.Add((eChartType)chartType);
    }
    #endregion
    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count
    {
        get
        {
            return this._list.Count;
        }
    }
    IEnumerator<ExcelChart> IEnumerable<ExcelChart>.GetEnumerator()
    {
        return this._list.GetEnumerator();
    }
    IEnumerator IEnumerable.GetEnumerator()
    {
        return this._list.GetEnumerator();
    }
    /// <summary>
    /// Returns a chart at the specific position.  
    /// </summary>
    /// <param name="PositionID">The position of the chart. 0-base</param>
    /// <returns></returns>
    public ExcelChart this[int PositionID]
    {
        get
        {
            return this._list[PositionID];
        }
    }
}