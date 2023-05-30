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
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// A serie for a pie chart
/// </summary>
public sealed class ExcelPieChartSerie : ExcelChartStandardSerie, IDrawingSerieDataLabel, IDrawingChartDataPoints
{
    /// <summary>
    /// Default constructor
    /// </summary>
    /// <param name="chart">The chart</param>
    /// <param name="ns">Namespacemanager</param>
    /// <param name="node">Topnode</param>
    /// <param name="isPivot">Is pivotchart</param>
    internal ExcelPieChartSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot)
        : base(chart, ns, node, isPivot)
    {
    }

    const string explosionPath = "c:explosion/@val";

    /// <summary>
    /// Explosion for Piecharts
    /// </summary>
    public int Explosion
    {
        get { return this.GetXmlNodeInt(explosionPath); }
        set
        {
            if (value < 0 || value > 400)
            {
                throw new ArgumentOutOfRangeException("Explosion range is 0-400");
            }

            this.SetXmlNodeString(explosionPath, value.ToString());
        }
    }

    ExcelChartSerieDataLabel _DataLabel;

    /// <summary>
    /// DataLabels
    /// </summary>
    public ExcelChartSerieDataLabel DataLabel
    {
        get { return this._DataLabel ??= new ExcelChartSerieDataLabel(this._chart, this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder); }
    }

    /// <summary>
    /// If the chart has datalabel
    /// </summary>
    public bool HasDataLabel
    {
        get { return this.TopNode.SelectSingleNode("c:dLbls", this.NameSpaceManager) != null; }
    }

    ExcelChartDataPointCollection _dataPoints;

    /// <summary>
    /// A collection of the individual datapoints
    /// </summary>
    public ExcelChartDataPointCollection DataPoints
    {
        get { return this._dataPoints ??= new ExcelChartDataPointCollection(this._chart, this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder); }
    }
}