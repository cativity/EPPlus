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
using System.Xml;
using OfficeOpenXml.Core;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// A chart legend
/// </summary>
public class ExcelChartLegend : XmlHelper, IDrawingStyle, IStyleMandatoryProperties
{
    internal ExcelChart _chart;
    internal string _nsPrefix;
    private readonly string OVERLAY_PATH;

    internal ExcelChartLegend(XmlNamespaceManager ns, XmlNode node, ExcelChart chart, string nsPrefix)
        : base(ns, node)
    {
        this._chart = chart;
        this._nsPrefix = nsPrefix;

        if (chart._isChartEx)
        {
            this.OVERLAY_PATH = "@overlay";
        }
        else
        {
            this.OVERLAY_PATH = "c:overlay/@val";
        }

        this.AddSchemaNodeOrder(new string[] { "legendPos", "legendEntry", "layout", "overlay", "spPr", "txPr" }, ExcelDrawing._schemaNodeOrderSpPr);
        _ = this.LoadLegendEntries();
    }

    internal void LoadEntries()
    {
        if (this._chart._isChartEx)
        {
            return;
        }

        this._entries = new EPPlusReadOnlyList<ExcelChartLegendEntry>();
        List<ExcelChartLegendEntry>? e = this.LoadLegendEntries();

        foreach (ExcelChart? c in this._chart.PlotArea.ChartTypes)
        {
            for (int i = 0; i < this._chart.Series.Count; i++)
            {
                int ix = e.FindIndex(x => x.Index == i);

                if (ix >= 0)
                {
                    this._entries.Add(e[ix]);
                }
                else
                {
                    this.AddNewEntry(this._chart.Series[i]);
                }
            }
        }
    }

    internal void AddNewEntry(ExcelChartSerie serie)
    {
        ExcelAddressBase? a = new ExcelAddressBase(serie.Series);

        if (a.Rows < 1 || a.Columns < 1)
        {
            return;
        }

        int seriesCount = a.Rows == 1 ? a.Rows : a.Columns;

        for (int i = 0; i < seriesCount; i++)
        {
            ExcelChartLegendEntry? entry = new ExcelChartLegendEntry(this.NameSpaceManager, this.TopNode, (ExcelChartStandard)this._chart, this._entries.Count);
            this._entries.Add(entry);
        }
    }

    internal int GetPreEntryIndex(int serieIndex)
    {
        for (int i = 0; i < this.Entries.Count; i++)
        {
            if (this.Entries[i].Index > serieIndex && this.Entries[i].TopNode.LocalName == "legendEntry")
            {
                return i;
            }
        }

        return -1;
    }

    internal EPPlusReadOnlyList<ExcelChartLegendEntry> _entries;

    /// <summary>
    /// A list of individual settings for legend entries.
    /// </summary>
    public EPPlusReadOnlyList<ExcelChartLegendEntry> Entries
    {
        get
        {
            if (this._entries == null)
            {
                this.LoadEntries();
            }

            return this._entries;
        }
    }

    internal XmlElement GetOrCreateEntry() => (XmlElement)this.CreateNode("c:legendEntry");

    internal List<ExcelChartLegendEntry> LoadLegendEntries()
    {
        if (this is ExcelChartExLegend)
        {
            return new List<ExcelChartLegendEntry>(); //Legend entries are not applicable for extended charts.
        }

        List<ExcelChartLegendEntry>? entries = new List<ExcelChartLegendEntry>();
        XmlNodeList? nodes = this.GetNodes("c:legendEntry");

        foreach (XmlNode n in nodes)
        {
            entries.Add(new ExcelChartLegendEntry(this.NameSpaceManager, n, (ExcelChartStandard)this._chart));
        }

        return entries;
    }

    const string POSITION_PATH = "c:legendPos/@val";

    /// <summary>
    /// The position of the Legend
    /// </summary>
    public virtual eLegendPosition Position
    {
        get
        {
            switch (this.GetXmlNodeString(POSITION_PATH).ToLower(CultureInfo.InvariantCulture))
            {
                case "t":
                    return eLegendPosition.Top;

                case "b":
                    return eLegendPosition.Bottom;

                case "l":
                    return eLegendPosition.Left;

                case "tr":
                    return eLegendPosition.TopRight;

                default:
                    return eLegendPosition.Right;
            }
        }
        set
        {
            if (this.TopNode == null)
            {
                throw new Exception("Can't set position. Chart has no legend");
            }

            switch (value)
            {
                case eLegendPosition.Top:
                    this.SetXmlNodeString(POSITION_PATH, "t");

                    break;

                case eLegendPosition.Bottom:
                    this.SetXmlNodeString(POSITION_PATH, "b");

                    break;

                case eLegendPosition.Left:
                    this.SetXmlNodeString(POSITION_PATH, "l");

                    break;

                case eLegendPosition.TopRight:
                    this.SetXmlNodeString(POSITION_PATH, "tr");

                    break;

                default:
                    this.SetXmlNodeString(POSITION_PATH, "r");

                    break;
            }
        }
    }

    /// <summary>
    /// If the legend overlays other objects
    /// </summary>
    public virtual bool Overlay
    {
        get => this.GetXmlNodeBool(this.OVERLAY_PATH);
        set
        {
            if (this.TopNode == null)
            {
                throw new Exception("Can't set overlay. Chart has no legend");
            }

            this.SetXmlNodeBool(this.OVERLAY_PATH, value);
        }
    }

    ExcelDrawingFill _fill;

    /// <summary>
    /// The Fill style
    /// </summary>
    public ExcelDrawingFill Fill => this._fill ??= new ExcelDrawingFill(this._chart, this.NameSpaceManager, this.TopNode, "c:spPr", this.SchemaNodeOrder);

    ExcelDrawingBorder _border;

    /// <summary>
    /// The Border style
    /// </summary>
    public ExcelDrawingBorder Border =>
        this._border ??= new ExcelDrawingBorder(this._chart,
                                                this.NameSpaceManager,
                                                this.TopNode,
                                                $"{this._nsPrefix}:spPr/a:ln",
                                                this.SchemaNodeOrder);

    ExcelTextFont _font;

    /// <summary>
    /// The Font properties
    /// </summary>
    public ExcelTextFont Font =>
        this._font ??= new ExcelTextFont(this._chart,
                                         this.NameSpaceManager,
                                         this.TopNode,
                                         $"{this._nsPrefix}:txPr/a:p/a:pPr/a:defRPr",
                                         this.SchemaNodeOrder);

    ExcelTextBody _textBody;

    /// <summary>
    /// Access to text body properties
    /// </summary>
    public ExcelTextBody TextBody => this._textBody ??= new ExcelTextBody(this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:txPr/a:bodyPr", this.SchemaNodeOrder);

    ExcelDrawingEffectStyle _effect;

    /// <summary>
    /// Effects
    /// </summary>
    public ExcelDrawingEffectStyle Effect =>
        this._effect ??= new ExcelDrawingEffectStyle(this._chart,
                                                     this.NameSpaceManager,
                                                     this.TopNode,
                                                     $"{this._nsPrefix}:spPr/a:effectLst",
                                                     this.SchemaNodeOrder);

    ExcelDrawing3D _threeD;

    /// <summary>
    /// 3D properties
    /// </summary>
    public ExcelDrawing3D ThreeD => this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager, this.TopNode, $"{this._nsPrefix}:spPr", this.SchemaNodeOrder);

    void IDrawingStyleBase.CreatespPr() => this.CreatespPrNode($"{this._nsPrefix}:spPr");

    /// <summary>
    /// Remove the legend
    /// </summary>
    public void Remove()
    {
        if (this.TopNode == null)
        {
            return;
        }

        _ = this.TopNode.ParentNode.RemoveChild(this.TopNode);
        this.TopNode = null;
    }

    /// <summary>
    /// Adds a legend to the chart
    /// </summary>
    public virtual void Add()
    {
        if (this.TopNode != null)
        {
            return;
        }

        //XmlHelper xml = new XmlHelper(NameSpaceManager, _chart.ChartXml);
        XmlHelper xml = XmlHelperFactory.Create(this.NameSpaceManager, this._chart.ChartXml);
        xml.SchemaNodeOrder = this._chart.SchemaNodeOrder;

        _ = xml.CreateNode("c:chartSpace/c:chart/c:legend");
        this.TopNode = this._chart.ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", this.NameSpaceManager);
        this.TopNode.InnerXml = "<c:legendPos val=\"r\" /><c:layout /><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>";
    }

    void IStyleMandatoryProperties.SetMandatoryProperties()
    {
        this.TextBody.Anchor = eTextAnchoringType.Center;
        this.TextBody.AnchorCenter = true;
        this.TextBody.WrapText = eTextWrappingType.Square;
        this.TextBody.VerticalTextOverflow = eTextVerticalOverflow.Ellipsis;
        this.TextBody.ParagraphSpacing = true;
        this.TextBody.Rotation = 0;

        if (this.Font.Kerning == 0)
        {
            this.Font.Kerning = 12;
        }

        this.Font.Bold = this.Font.Bold; //Must be set

        this.CreatespPrNode($"{this._nsPrefix}:spPr");
    }
}