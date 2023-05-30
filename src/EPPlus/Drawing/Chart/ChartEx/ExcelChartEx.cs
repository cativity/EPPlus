/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/

using OfficeOpenXml.Constants;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx;

/// <summary>
/// Base class for all extention charts
/// </summary>
public abstract class ExcelChartEx : ExcelChart
{
    internal ExcelChartEx(ExcelDrawings drawings, XmlNode node, ExcelGroupShape parent)
        : base(drawings, node, parent, "mc:AlternateContent/mc:Choice/xdr:graphicFrame")
    {
        this.ChartType = GetChartType(node, drawings.NameSpaceManager);
        this.Init();
    }

    internal ExcelChartEx(ExcelDrawings drawings, XmlNode drawingsNode, eChartType? type, XmlDocument chartXml = null, ExcelGroupShape parent = null)
        : base(drawings, drawingsNode, chartXml, parent, "mc:AlternateContent/mc:Choice/xdr:graphicFrame")
    {
        this.ChartType = type.Value;
        this.CreateNewChart(drawings, chartXml, type);
        this.Init();
    }

    internal ExcelChartEx(ExcelDrawings drawings,
                          XmlNode node,
                          Uri uriChart,
                          ZipPackagePart part,
                          XmlDocument chartXml,
                          XmlNode chartNode,
                          ExcelGroupShape parent = null)
        : base(drawings, node, chartXml, parent, "mc:AlternateContent/mc:Choice/xdr:graphicFrame")
    {
        this.UriChart = uriChart;
        this.Part = part;
        this._chartNode = chartNode;
        this._chartXmlHelper = XmlHelperFactory.Create(drawings.NameSpaceManager, chartNode);
        this.ChartType = GetChartType(chartNode, drawings.NameSpaceManager);
        this.Init();
    }

    internal void LoadAxis()
    {
        List<ExcelChartAxis>? l = new List<ExcelChartAxis>();

        foreach (XmlNode axNode in this._chartXmlHelper.GetNodes("cx:plotArea/cx:axis"))
        {
            l.Add(new ExcelChartExAxis(this, this.NameSpaceManager, axNode));
        }

        this._axis = l.ToArray();
        this._exAxis = null;

        if (this.Axis.Length > 0)
        {
            if (this.Axis[1].AxisType == eAxisType.Cat)
            {
                this.XAxis = this.Axis[1];
                this.YAxis = this.Axis[0];
            }
            else
            {
                this.XAxis = this.Axis[0];
                this.YAxis = this.Axis[1];
            }
        }
    }

    private void Init()
    {
        this._isChartEx = true;

        this._chartXmlHelper.SchemaNodeOrder = new string[]
        {
            "chartData", "chart", "spPr", "txPr", "clrMapOvr", "fmtOvrs", "title", "plotArea", "plotAreaRegion", "axis", "legend", "printSettings"
        };

        base.Series.Init(this, this.NameSpaceManager, this._chartNode, false);
        this.Series.Init(this, this.NameSpaceManager, this._chartNode, false, base.Series._list);
        this.LoadAxis();
    }

    private void CreateNewChart(ExcelDrawings drawings, XmlDocument chartXml = null, eChartType? type = null)
    {
        XmlElement graphFrame = this.TopNode.OwnerDocument.CreateElement("mc", "AlternateContent", ExcelPackage.schemaMarkupCompatibility);
        graphFrame.SetAttribute("xmlns:mc", ExcelPackage.schemaMarkupCompatibility);
        _ = this.TopNode.AppendChild(graphFrame);

        graphFrame.InnerXml =
            string.Format(
                          "<mc:Choice xmlns:cx1=\"{1}\" Requires=\"cx1\"><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"{0}\" name=\"\"><a:extLst><a:ext uri=\"{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}\"><a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{{9FE3C5B3-14FE-44E2-AB27-50960A44C7C4}}\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2014/chartex\"><cx:chart xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"{0}\" name=\"\"/><xdr:cNvSpPr><a:spLocks noTextEdit=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"3609974\" y=\"938212\"/><a:ext cx=\"5762625\" cy=\"2743200\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val=\"white\"/></a:solidFill><a:ln w=\"1\"><a:solidFill><a:prstClr val=\"green\"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\"/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\" sz=\"1100\"/><a:t>This chart isn't available in your version of Excel. Editing this shape or saving this workbook into a different file format will permanently break the chart.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>",
                          this._id,
                          GetChartExNameSpace(type ?? eChartType.Sunburst));

        _ = this.TopNode.AppendChild(this.TopNode.OwnerDocument.CreateElement("clientData", ExcelPackage.schemaSheetDrawings));

        ZipPackage? package = drawings.Worksheet._package.ZipPackage;
        this.UriChart = GetNewUri(package, "/xl/charts/chartex{0}.xml");

        if (chartXml == null)
        {
            this.ChartXml = new XmlDocument { PreserveWhitespace = ExcelPackage.preserveWhitespace };
            LoadXmlSafe(this.ChartXml, ChartStartXml(type.Value), Encoding.UTF8);
        }
        else
        {
            this.ChartXml = chartXml;
        }

        // save it to the package
        this.Part = package.CreatePart(this.UriChart, ContentTypes.contentTypeChartEx, this._drawings._package.Compression);

        StreamWriter streamChart = new StreamWriter(this.Part.GetStream(FileMode.Create, FileAccess.Write));
        this.ChartXml.Save(streamChart);
        streamChart.Close();
        ZipPackage.Flush();

        ZipPackageRelationship? chartRelation = drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(drawings.UriDrawing, this.UriChart),
                                                                                 TargetMode.Internal,
                                                                                 ExcelPackage.schemaChartExRelationships);

        graphFrame.SelectSingleNode("mc:Choice/xdr:graphicFrame/a:graphic/a:graphicData/cx:chart", this.NameSpaceManager).Attributes["r:id"].Value =
            chartRelation.Id;

        ZipPackage.Flush();
        this._chartNode = this.ChartXml.SelectSingleNode("cx:chartSpace/cx:chart", this.NameSpaceManager);
        this._chartXmlHelper = XmlHelperFactory.Create(this.NameSpaceManager, this._chartNode);
        this.GetPositionSize();
    }

    private static string GetChartExNameSpace(eChartType type)
    {
        switch (type)
        {
            case eChartType.RegionMap:
                return "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex";

            case eChartType.Funnel:
                return "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex";

            default:
                return "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex";
        }
    }

    private static string ChartStartXml(eChartType type)
    {
        StringBuilder xml = new StringBuilder();

        _ = xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        _ = xml.Append("<cx:chartSpace xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" >");
        _ = xml.Append("<cx:chart><cx:title overlay=\"0\" align=\"ctr\" pos=\"t\"/><cx:plotArea><cx:plotAreaRegion></cx:plotAreaRegion></cx:plotArea></cx:chart>");
        _ = xml.Append("</cx:chartSpace>");

        return xml.ToString();
    }

    private static void AddData(StringBuilder xml)
    {
        _ = xml.Append("<cx:chartData><cx:data id=\"0\"><cx:strDim type=\"cat\"><cx:f dir=\"row\">_xlchart.v1.31</cx:f></cx:strDim><cx:numDim type=\"size\"><cx:f dir=\"row\">_xlchart.v1.32</cx:f></cx:numDim></cx:data><cx:data id=\"1\"><cx:strDim type=\"cat\"><cx:f dir=\"row\">_xlchart.v1.31</cx:f></cx:strDim><cx:numDim type=\"size\"><cx:f dir=\"row\">_xlchart.v1.33</cx:f></cx:numDim></cx:data></cx:chartData>");
    }

    internal override void AddAxis()
    {
        List<ExcelChartAxis>? l = new List<ExcelChartAxis>();

        foreach (XmlNode axNode in this._chartXmlHelper.GetNodes("cx:plotArea/cx:axis"))
        {
            l.Add(new ExcelChartExAxis(this, this.NameSpaceManager, axNode));
        }
    }

    private static eChartType GetChartType(XmlNode node, XmlNamespaceManager nsm)
    {
        XmlNode? layoutId = node.SelectSingleNode("cx:plotArea/cx:plotAreaRegion/cx:series[1]/@layoutId", nsm);

        if (layoutId == null)
        {
            return eChartType.Treemap;
        }

        switch (layoutId.Value)
        {
            case "clusteredColumn":
                layoutId = node.SelectSingleNode("cx:plotArea/cx:plotAreaRegion/cx:series[@layoutId='paretoLine']", nsm);

                if (layoutId == null)
                {
                    return eChartType.Histogram;
                }
                else
                {
                    return eChartType.Pareto;
                }

            case "paretoLine":
                return eChartType.Pareto;

            case "boxWhisker":
                return eChartType.BoxWhisker;

            case "funnel":
                return eChartType.Funnel;

            case "regionMap":
                return eChartType.RegionMap;

            case "sunburst":
                return eChartType.Sunburst;

            case "treemap":
                return eChartType.Treemap;

            case "waterfall":
                return eChartType.Waterfall;

            default:
                throw new InvalidOperationException($"Unsupported layoutId in ChartEx Xml: {layoutId}");
        }
    }

    /// <summary>
    /// Delete the charts title
    /// </summary>
    public override void DeleteTitle()
    {
        this._chartXmlHelper.DeleteNode("cx:title");
    }

    /// <summary>
    /// Plotarea properties
    /// </summary>
    public override ExcelChartPlotArea PlotArea
    {
        get
        {
            if (this._plotArea == null)
            {
                XmlNode? node = this._chartXmlHelper.GetNode("cx:plotArea");
                this._plotArea = new ExcelChartExPlotarea(this.NameSpaceManager, node, this);
            }

            return this._plotArea;
        }
    }

    internal ExcelChartExAxis[] _exAxis;

    /// <summary>
    /// An array containg all axis of all Charttypes
    /// </summary>
    public new ExcelChartExAxis[] Axis
    {
        get { return this._exAxis ??= this._axis.Select(x => (ExcelChartExAxis)x).ToArray(); }
    }

    /// <summary>
    /// The titel of the chart
    /// </summary>
    public new ExcelChartExTitle Title
    {
        get
        {
            this._title ??= this.GetTitle();

            return (ExcelChartExTitle)this._title;
        }
    }

    internal override ExcelChartTitle GetTitle()
    {
        return new ExcelChartExTitle(this, this.NameSpaceManager, this.ChartXml.SelectSingleNode("cx:chartSpace/cx:chart", this.NameSpaceManager));
    }

    /// <summary>
    /// Legend
    /// </summary>
    public new ExcelChartExLegend Legend
    {
        get
        {
            if (this._legend == null)
            {
                return (ExcelChartExLegend)base.Legend;
            }

            return (ExcelChartExLegend)this._legend;
        }
    }

    ExcelDrawingBorder _border;

    /// <summary>
    /// Border
    /// </summary>
    public override ExcelDrawingBorder Border
    {
        get
        {
            return this._border ??= new ExcelDrawingBorder(this,
                                                           this.NameSpaceManager,
                                                           this.ChartXml.SelectSingleNode("cx:chartSpace", this.NameSpaceManager),
                                                           "cx:spPr/a:ln",
                                                           this._chartXmlHelper.SchemaNodeOrder);
        }
    }

    ExcelDrawingFill _fill;

    /// <summary>
    /// Access to Fill properties
    /// </summary>
    public override ExcelDrawingFill Fill
    {
        get
        {
            return this._fill ??= new ExcelDrawingFill(this,
                                                       this.NameSpaceManager,
                                                       this.ChartXml.SelectSingleNode("cx:chartSpace", this.NameSpaceManager),
                                                       "cx:spPr",
                                                       this._chartXmlHelper.SchemaNodeOrder);
        }
    }

    ExcelDrawingEffectStyle _effect;

    /// <summary>
    /// Effects
    /// </summary>
    public override ExcelDrawingEffectStyle Effect
    {
        get
        {
            return this._effect ??= new ExcelDrawingEffectStyle(this,
                                                                this.NameSpaceManager,
                                                                this.ChartXml.SelectSingleNode("cx:chartSpace", this.NameSpaceManager),
                                                                "cx:spPr/a:effectLst",
                                                                this._chartXmlHelper.SchemaNodeOrder);
        }
    }

    ExcelDrawing3D _threeD;

    /// <summary>
    /// 3D properties
    /// </summary>
    public override ExcelDrawing3D ThreeD
    {
        get
        {
            return this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager,
                                                       this.ChartXml.SelectSingleNode("cx:chartSpace", this.NameSpaceManager),
                                                       "cx:spPr",
                                                       this._chartXmlHelper.SchemaNodeOrder);
        }
    }

    ExcelTextFont _font;

    /// <summary>
    /// Access to font properties
    /// </summary>
    public override ExcelTextFont Font
    {
        get
        {
            return this._font ??= new ExcelTextFont(this,
                                                    this.NameSpaceManager,
                                                    this.ChartXml.SelectSingleNode("cx:chartSpace", this.NameSpaceManager),
                                                    "cx:txPr/a:p/a:pPr/a:defRPr",
                                                    this._chartXmlHelper.SchemaNodeOrder);
        }
    }

    ExcelTextBody _textBody;

    /// <summary>
    /// Access to text body properties
    /// </summary>
    public override ExcelTextBody TextBody
    {
        get
        {
            return this._textBody ??= new ExcelTextBody(this.NameSpaceManager,
                                                        this.ChartXml.SelectSingleNode("cx:chartSpace", this.NameSpaceManager),
                                                        "cx:txPr/a:bodyPr",
                                                        this._chartXmlHelper.SchemaNodeOrder);
        }
    }

    /// <summary>
    /// Chart series
    /// </summary>
    public new ExcelChartSeries<ExcelChartExSerie> Series { get; } = new ExcelChartSeries<ExcelChartExSerie>();

    /// <summary>
    /// Is not applied to Extension charts
    /// </summary>
    public override bool VaryColors
    {
        get { return false; }
        set { throw new InvalidOperationException("VaryColors do not apply to Extended charts"); }
    }

    /// <summary>
    /// Cannot be set for extension charts. Please use <see cref="ExcelChart.StyleManager"/>
    /// </summary>
    public override eChartStyle Style { get; set; }

    /// <summary>
    /// If the chart has a title or not
    /// </summary>
    public override bool HasTitle
    {
        get { return this._chartXmlHelper.ExistsNode("cx:title"); }
    }

    /// <summary>
    /// If the chart has legend or not
    /// </summary>
    public override bool HasLegend
    {
        get { return this._chartXmlHelper.ExistsNode("cx:legend"); }
    }

    /// <summary>
    /// 3D settings
    /// </summary>
    public override ExcelView3D View3D
    {
        get { return null; }
    }

    /// <summary>
    /// This property does not apply to extended charts.
    /// This property will always return eDisplayBlanksAs.Zero.
    /// Setting this property on an extended chart will result in an InvalidOperationException
    /// </summary>
    public override eDisplayBlanksAs DisplayBlanksAs
    {
        get { return eDisplayBlanksAs.Zero; }
        set { throw new InvalidOperationException("DisplayBlanksAs do not apply to Extended charts"); }
    }

    /// <summary>
    /// This property does not apply to extended charts.
    /// Setting this property on an extended chart will result in an InvalidOperationException
    /// </summary>
    public override bool RoundedCorners
    {
        get { return false; }
        set { throw new InvalidOperationException("RoundedCorners do not apply to Extended charts"); }
    }

    /// <summary>
    /// This property does not apply to extended charts.
    /// Setting this property on an extended chart will result in an InvalidOperationException
    /// </summary>
    public override bool ShowDataLabelsOverMaximum
    {
        get { return false; }
        set { throw new InvalidOperationException("ShowHiddenData do not apply to Extended charts"); }
    }

    /// <summary>
    /// This property does not apply to extended charts.
    /// Setting this property on an extended chart will result in an InvalidOperationException
    /// </summary>
    public override bool ShowHiddenData
    {
        get { return false; }
        set { throw new InvalidOperationException("ShowHiddenData do not apply to Extended charts"); }
    }

    /// <summary>
    /// The X Axis
    /// </summary>
    public new ExcelChartExAxis XAxis
    {
        get { return (ExcelChartExAxis)base.XAxis; }
        internal set { base.XAxis = value; }
    }

    /// <summary>
    /// The Y Axis
    /// </summary>
    public new ExcelChartExAxis YAxis
    {
        get { return (ExcelChartExAxis)base.YAxis; }
        internal set { base.YAxis = value; }
    }
}