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
/// A series for an Area Chart
/// </summary>
public sealed class ExcelAreaChartSerie : ExcelChartSerieWithHorizontalErrorBars, IDrawingSerieDataLabel, IDrawingChartDataPoints
{
    /// <summary>
    /// Default constructor
    /// </summary>
    /// <param name="chart">Chart series</param>
    /// <param name="ns">Namespacemanager</param>
    /// <param name="node">Topnode</param>
    /// <param name="isPivot">Is pivotchart</param>
    internal ExcelAreaChartSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot)
        : base(chart, ns, node, isPivot)
    {
    }

    ExcelChartSerieDataLabel _DataLabel;

    /// <summary>
    /// Datalabel
    /// </summary>
    public ExcelChartSerieDataLabel DataLabel => this._DataLabel ??= new ExcelChartSerieDataLabel(this._chart, this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder);

    /// <summary>
    /// If the chart has datalabel
    /// </summary>
    public bool HasDataLabel => this.TopNode.SelectSingleNode("c:dLbls", this.NameSpaceManager) != null;

    const string INVERTIFNEGATIVE_PATH = "c:invertIfNegative/@val";

    internal bool InvertIfNegative
    {
        get => this.GetXmlNodeBool(INVERTIFNEGATIVE_PATH, true);
        set => this.SetXmlNodeBool(INVERTIFNEGATIVE_PATH, value);
    }

    ExcelChartDataPointCollection _dataPoints;

    /// <summary>
    /// A collection of the individual datapoints
    /// </summary>
    public ExcelChartDataPointCollection DataPoints => this._dataPoints ??= new ExcelChartDataPointCollection(this._chart, this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder);
}