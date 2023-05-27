using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx;

/// <summary>
/// A legend for an Extended chart
/// </summary>
public class ExcelChartExLegend : ExcelChartLegend
{
    internal ExcelChartExLegend(ExcelChart chart, XmlNamespaceManager nsm, XmlNode node)
        : base(nsm, node, chart, "cx")
    {
        this.SchemaNodeOrder = new string[] { "spPr", "txPr" };
    }

    /// <summary>
    /// The side position alignment of the legend
    /// </summary>
    public ePositionAlign PositionAlignment
    {
        get { return this.GetXmlNodeString("@align").Replace("ctr", "center").ToEnum(ePositionAlign.Center); }
        set
        {
            if (this.TopNode == null)
            {
                this.Add();
            }

            this.SetXmlNodeString("@align", value.ToEnumString().Replace("center", "ctr"));
        }
    }

    /// <summary>
    /// The position of the Legend.
    /// </summary>
    /// <remarks>Setting the Position to TopRight will set the <see cref="Position"/> to Right and the <see cref="PositionAlignment" /> to Min</remarks>
    public override eLegendPosition Position
    {
        get
        {
            switch (this.GetXmlNodeString("@pos"))
            {
                case "l":
                    return eLegendPosition.Left;

                case "r":
                    return eLegendPosition.Right;

                case "b":
                    return eLegendPosition.Bottom;

                default:
                    return eLegendPosition.Top;
            }
        }
        set
        {
            if (this.TopNode == null)
            {
                this.Add();
            }

            if (value == eLegendPosition.TopRight)
            {
                this.PositionAlignment = ePositionAlign.Min;
                value = eLegendPosition.Right;
            }

            this.SetXmlNodeString("@pos", value.ToEnumString().Substring(0, 1).ToLowerInvariant());
        }
    }

    /// <summary>
    /// Adds a legend to the chart
    /// </summary>
    public override void Add()
    {
        if (this.TopNode != null)
        {
            return;
        }

        //XmlHelper xml = new XmlHelper(NameSpaceManager, _chart.ChartXml);
        XmlHelper xml = XmlHelperFactory.Create(this.NameSpaceManager, this._chart.ChartXml);
        xml.SchemaNodeOrder = this._chart.SchemaNodeOrder;

        this.TopNode = xml.CreateNode("cx:chartSpace/cx:chart/cx:legend");
    }
}