/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB           EPPlus 5.2
 *************************************************************************************************/

using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx;

/// <summary>
/// A series for an Histogram Chart
/// </summary>
public class ExcelHistogramChartSerie : ExcelChartExSerie
{
    internal int _index;

    internal ExcelHistogramChartSerie(ExcelChartEx chart, XmlNamespaceManager ns, XmlNode node, int index = -1)
        : base(chart, ns, node)
    {
        if (index == -1)
        {
            this._index = chart.Series.Count * (chart.ChartType == eChartType.Pareto ? 2 : 1);
        }
        else
        {
            this._index = index;
        }
    }

    internal void AddParetoLine()
    {
        int ix = this._chart.Series.Count * 2;
        XmlElement? serElement = CreateSeriesElement((ExcelChartEx)this._chart, eChartType.Pareto, ix + 1, this.TopNode, true);
        serElement.SetAttribute("ownerIdx", ix.ToString());
        serElement.InnerXml = "<cx:axisId val=\"2\"/>";
        this.AddParetoLineFromSerie(serElement);
    }

    ExcelChartExSerieBinning _binning;

    /// <summary>
    /// The data binning properties
    /// </summary>
    public ExcelChartExSerieBinning Binning => this._binning ??= new ExcelChartExSerieBinning(this.NameSpaceManager, this.TopNode);

    internal const string _aggregationPath = "cx:layoutPr/cx:aggregation";
    internal const string _binningPath = "cx:layoutPr/cx:binning";

    /// <summary>
    /// If x-axis is per category
    /// </summary>
    public bool Aggregation
    {
        get => this.ExistsNode(_aggregationPath);
        set
        {
            if (value)
            {
                this.DeleteNode(_binningPath);
                _ = this.CreateNode(_aggregationPath);
            }
            else
            {
                this.DeleteNode(_aggregationPath);

                if (!this.ExistsNode(_binningPath))
                {
                    this.Binning.IntervalClosed = eIntervalClosed.Right;
                }
            }
        }
    }

    internal void AddParetoLineFromSerie(XmlElement serElement) => this.ParetoLine = new ExcelChartExParetoLine(this._chart, this.NameSpaceManager, serElement);

    internal void RemoveParetoLine()
    {
        this.ParetoLine?.DeleteNode(".");
        this.ParetoLine = null;
    }

    /// <summary>
    /// Properties for the pareto line.
    /// </summary>
    public ExcelChartExParetoLine ParetoLine { get; private set; }
}