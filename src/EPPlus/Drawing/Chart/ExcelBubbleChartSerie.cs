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
/// A serie for a bubble chart
/// </summary>
public sealed class ExcelBubbleChartSerie : ExcelChartSerieWithHorizontalErrorBars, IDrawingSerieDataLabel, IDrawingChartDataPoints
{
    /// <summary>
    /// Default constructor
    /// </summary>
    /// <param name="chart">The chart</param>
    /// <param name="ns">Namespacemanager</param>
    /// <param name="node">Topnode</param>
    /// <param name="isPivot">Is pivotchart</param>
    internal ExcelBubbleChartSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot)
        : base(chart, ns, node, isPivot)
    {
    }

    ExcelChartSerieDataLabel _dataLabel = null;

    /// <summary>
    /// Datalabel
    /// </summary>
    public ExcelChartSerieDataLabel DataLabel
    {
        get { return this._dataLabel ??= new ExcelChartSerieDataLabel(this._chart, this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder); }
    }

    /// <summary>
    /// If the chart has datalabel
    /// </summary>
    public bool HasDataLabel
    {
        get { return this.TopNode.SelectSingleNode("c:dLbls", this.NameSpaceManager) != null; }
    }

    const string BUBBLE3D_PATH = "c:bubble3D/@val";

    internal bool Bubble3D
    {
        get { return this.GetXmlNodeBool(BUBBLE3D_PATH, true); }
        set { this.SetXmlNodeBool(BUBBLE3D_PATH, value); }
    }

    const string INVERTIFNEGATIVE_PATH = "c:invertIfNegative/@val";

    internal bool InvertIfNegative
    {
        get { return this.GetXmlNodeBool(INVERTIFNEGATIVE_PATH, true); }
        set { this.SetXmlNodeBool(INVERTIFNEGATIVE_PATH, value); }
    }

    /// <summary>
    /// The dataseries for the Bubble Chart
    /// </summary>
    public override string Series
    {
        get { return base.Series; }
        set
        {
            base.Series = value;

            if (string.IsNullOrEmpty(this.BubbleSize))
            {
                this.GenerateLit();
            }
        }
    }

    const string BUBBLESIZE_TOPPATH = "c:bubbleSize";
    const string BUBBLESIZE_PATH = BUBBLESIZE_TOPPATH + "/c:numRef/c:f";

    /// <summary>
    /// The size of the bubbles
    /// </summary>
    public string BubbleSize
    {
        get { return this.GetXmlNodeString(BUBBLESIZE_PATH); }
        set
        {
            if (string.IsNullOrEmpty(value))
            {
                this.GenerateLit();
            }
            else
            {
                this.SetXmlNodeString(BUBBLESIZE_PATH, ExcelCellBase.GetFullAddress(this._chart.WorkSheet.Name, value));

                XmlNode cache = this.TopNode.SelectSingleNode(string.Format("{0}/c:numCache", BUBBLESIZE_PATH), this.NameSpaceManager);

                if (cache != null)
                {
                    _ = cache.ParentNode.RemoveChild(cache);
                }

                this.DeleteNode(string.Format("{0}/c:numLit", BUBBLESIZE_TOPPATH));
            }
        }
    }

    internal void GenerateLit()
    {
        ExcelAddress? s = new ExcelAddress(this.Series);
        int ix = 0;
        StringBuilder? sb = new StringBuilder();

        for (int row = s._fromRow; row <= s._toRow; row++)
        {
            for (int c = s._fromCol; c <= s._toCol; c++)
            {
                _ = sb.AppendFormat("<c:pt idx=\"{0}\"><c:v>1</c:v></c:pt>", ix++);
            }
        }

        _ = this.CreateNode(BUBBLESIZE_TOPPATH + "/c:numLit", true);
        XmlNode lit = this.TopNode.SelectSingleNode(string.Format("{0}/c:numLit", BUBBLESIZE_TOPPATH), this.NameSpaceManager);
        lit.InnerXml = string.Format("<c:formatCode>General</c:formatCode><c:ptCount val=\"{0}\"/>{1}", ix, sb.ToString());
    }

    ExcelChartDataPointCollection _dataPoints = null;

    /// <summary>
    /// A collection of the individual datapoints
    /// </summary>
    public ExcelChartDataPointCollection DataPoints
    {
        get { return this._dataPoints ??= new ExcelChartDataPointCollection(this._chart, this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder); }
    }
}