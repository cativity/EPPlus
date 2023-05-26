/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/15/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using System.IO;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Constants;
using System.Linq;
namespace OfficeOpenXml.Drawing.Chart;

/// <summary>
/// Base class for Chart object.
/// </summary>
public class ExcelChartStandard : ExcelChart
{
    #region "Constructors"
    internal ExcelChartStandard(ExcelDrawings drawings, XmlNode node, eChartType? type, bool isPivot, ExcelGroupShape parent, string drawingPath = "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
        base(drawings, node, parent, drawingPath, nvPrPath)
    {
        if (type.HasValue)
        {
            this.ChartType = type.Value;
        }

        this.CreateNewChart(drawings, null, null, type);

        this.Init(drawings, this._chartNode);
        this.InitSeries(this, drawings.NameSpaceManager, this._chartNode, isPivot);
        this.SetTypeProperties();
        this.LoadAxis();
    }
    internal ExcelChartStandard(ExcelDrawings drawings, XmlNode drawingsNode, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml = null, ExcelGroupShape parent = null, string drawingPath = "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
        base(drawings, drawingsNode, chartXml, parent, drawingPath, nvPrPath)
    {
        if (type.HasValue)
        {
            this.ChartType = type.Value;
        }

        this._topChart = topChart;
        this.CreateNewChart(drawings, topChart, chartXml, type);

        this.Init(drawings, this._chartNode);

        if (chartXml == null)
        {
            this.SetTypeProperties();
        }
        else
        {
            this.ChartType = this.GetChartType(this._chartNode.LocalName);
        }

        this.InitSeries(this, drawings.NameSpaceManager, this._chartNode, PivotTableSource != null);
        if (PivotTableSource != null)
        {
            this.SetPivotSource(PivotTableSource);
        }

        if (topChart == null)
        {
            this.LoadAxis();
        }
        else
        {
            this._axis = topChart.Axis;
            if (this._axis.Length > 0)
            {
                this.XAxis = (ExcelChartAxisStandard)this._axis[0];
                this.YAxis = (ExcelChartAxisStandard)this._axis[1];
            }
        }
    }
    internal ExcelChartStandard(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent, string drawingPath = "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
        base(drawings, node, chartXml, parent, drawingPath, nvPrPath)
    {
        string? ptSource = this._chartXmlHelper.GetXmlNodeString("c:pivotSource/c:name");
        if(!string.IsNullOrEmpty(ptSource))
        {
            if(ptSource.StartsWith("["))
            {
                ptSource = ptSource.Substring(ptSource.IndexOf("]") + 1);
            }
            string? wsName = ExcelAddressBase.GetWorksheetPart(ptSource,"");
            ExcelWorksheet? ws = drawings.Worksheet.Workbook.Worksheets[wsName];
            if(ws!=null)
            {
                string? ptName = ptSource.Substring(ptSource.LastIndexOf("!")+1);
                this.PivotTableSource = ws.PivotTables[ptName];
                this._chartXmlHelper.SetXmlNodeString("c:pivotSource/c:name", "[]"+ptSource);
            }
        }

        this.UriChart = uriChart;
        this.Part = part;
        this.ChartXml = chartXml;
        this._chartNode = chartNode;
        this.InitSeries(this, drawings.NameSpaceManager, this._chartNode, this.PivotTableSource != null);
        this.InitChartLoad(drawings, chartNode);
        this.ChartType = this.GetChartType(chartNode.LocalName);
    }
    internal ExcelChartStandard(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent, string drawingPath = "xdr:graphicFrame", string nvPrPath = "xdr:nvGraphicFramePr/xdr:cNvPr") :
        base(topChart, chartNode, parent, drawingPath, nvPrPath)
    {
        this.UriChart = topChart.UriChart;
        this.Part = topChart.Part;
        this.ChartXml = topChart.ChartXml;
        this._plotArea = topChart.PlotArea;
        this._chartNode = chartNode;
        this.InitSeries(this, topChart._drawings.NameSpaceManager, this._chartNode, false);
        this.InitChartLoad(topChart._drawings, chartNode);
    }
    private void InitChartLoad(ExcelDrawings drawings, XmlNode chartNode)
    {
        bool isPivot = false;
        this.Init(drawings, chartNode);
        this.InitSeries(this, drawings.NameSpaceManager, this._chartNode, isPivot);
        this.LoadAxis();
    }
    internal virtual void InitSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list = null)
    {
        this.Series.Init(chart, ns, node, isPivot, list);
    }
    private void Init(ExcelDrawings drawings, XmlNode chartNode)
    {
        this._isChartEx = chartNode.NamespaceURI == ExcelPackage.schemaChartExMain;
        this._chartXmlHelper = XmlHelperFactory.Create(drawings.NameSpaceManager, chartNode);
        this._chartXmlHelper.AddSchemaNodeOrder(new string[] { "date1904", "lang", "roundedCorners", "AlternateContent", "style", "clrMapOvr", "pivotSource", "protection", "chart", "ofPieType", "title", "pivotFmt", "autoTitleDeleted", "view3D", "floor", "sideWall", "backWall", "plotArea", "wireframe", "barDir", "grouping", "scatterStyle", "radarStyle", "varyColors", "ser", "dLbls", "bubbleScale", "showNegBubbles", "firstSliceAng", "holeSize", "dropLines", "hiLowLines", "upDownBars", "marker", "smooth", "shape", "legend", "plotVisOnly", "dispBlanksAs", "gapWidth", "upBars", "downBars", "showDLblsOverMax", "overlap", "bandFmts", "axId", "spPr", "txPr", "printSettings" }, _schemaNodeOrderSpPr);
        this.WorkSheet = drawings.Worksheet;
    }
    #endregion
    #region "Private functions"
    private void SetTypeProperties()
    {
        /******* Grouping *******/
        if (this.IsTypeClustered())
        {
            this.Grouping = eGrouping.Clustered;
        }
        else if (this.IsTypeStacked())
        {
            this.Grouping = eGrouping.Stacked;
        }
        else if (this.IsTypePercentStacked())
        {
            this.Grouping = eGrouping.PercentStacked;
        }

        /***** 3D Perspective *****/
        if (this.IsType3D())
        {
            this.View3D.RotY = 20;
            this.View3D.Perspective = 30;    //Default to 30
            if (this.IsTypePieDoughnut())
            {
                this.View3D.RotX = 30;
            }
            else
            {
                this.View3D.RotX = 15;
            }
        }
    }
    private void Init3DProperties()
    {
        this.Floor = new ExcelChartSurface(this, this.NameSpaceManager, this.ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:floor", this.NameSpaceManager));
        this.BackWall = new ExcelChartSurface(this, this.NameSpaceManager, this.ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:backWall", this.NameSpaceManager));
        this.SideWall = new ExcelChartSurface(this, this.NameSpaceManager, this.ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:sideWall", this.NameSpaceManager));
    }
    private void CreateNewChart(ExcelDrawings drawings, ExcelChart topChart, XmlDocument chartXml = null, eChartType? type = null)
    {
        if (topChart == null)
        {
            XmlElement graphFrame = this.TopNode.OwnerDocument.CreateElement("graphicFrame", ExcelPackage.schemaSheetDrawings);
            graphFrame.SetAttribute("macro", "");
            this.TopNode.AppendChild(graphFrame);
            graphFrame.InnerXml = string.Format("<xdr:nvGraphicFramePr><xdr:cNvPr id=\"{0}\" name=\"Chart 1\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\" /> <a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\" />   </a:graphicData>  </a:graphic>", this._id);
            this.TopNode.AppendChild(this.TopNode.OwnerDocument.CreateElement("clientData", ExcelPackage.schemaSheetDrawings));

            ZipPackage? package = drawings.Worksheet._package.ZipPackage;
            this.UriChart = GetNewUri(package, "/xl/charts/chart{0}.xml");

            if (chartXml == null)
            {
                this.ChartXml = new XmlDocument
                {
                    PreserveWhitespace = ExcelPackage.preserveWhitespace
                };
                LoadXmlSafe(this.ChartXml, this.ChartStartXml(type.Value), Encoding.UTF8);
            }
            else
            {
                this.ChartXml = chartXml;
            }

            // save it to the package
            this.Part = package.CreatePart(this.UriChart, ContentTypes.contentTypeChart, this._drawings._package.Compression);

            StreamWriter streamChart = new StreamWriter(this.Part.GetStream(FileMode.Create, FileAccess.Write));
            this.ChartXml.Save(streamChart);
            streamChart.Close();
            ZipPackage.Flush();

            ZipPackageRelationship? chartRelation = drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(drawings.UriDrawing, this.UriChart), TargetMode.Internal, ExcelPackage.schemaRelationships + "/chart");
            graphFrame.SelectSingleNode("a:graphic/a:graphicData/c:chart", this.NameSpaceManager).Attributes["r:id"].Value = chartRelation.Id;
            ZipPackage.Flush();
            this._chartNode = this.ChartXml.SelectSingleNode(string.Format("c:chartSpace/c:chart/c:plotArea/{0}", this.GetChartNodeText()), this.NameSpaceManager);
        }
        else
        {
            this.ChartXml = topChart.ChartXml;
            this.Part = topChart.Part;
            this._plotArea = topChart.PlotArea;
            this.UriChart = topChart.UriChart;
            this._axis = topChart._axis;

            XmlNode preNode = this._plotArea.ChartTypes[this._plotArea.ChartTypes.Count - 1].ChartNode;
            this._chartNode = ((XmlDocument)this.ChartXml).CreateElement(this.GetChartNodeText(), ExcelPackage.schemaChart);
            preNode.ParentNode.InsertAfter(this._chartNode, preNode);
            if (topChart.Axis.Length == 0)
            {
                this.AddAxis();
            }
            string serieXML = this.GetChartSerieStartXml(type.Value, int.Parse(topChart.Axis[0].Id), int.Parse(topChart.Axis[1].Id), topChart.Axis.Length > 2 ? int.Parse(topChart.Axis[2].Id) : -1);
            this._chartNode.InnerXml = serieXML;
        }

        this.GetPositionSize();
        if (this.IsType3D())
        {
            this.Init3DProperties();
        }
    }
    private void LoadAxis()
    {
        List<ExcelChartAxis> l = new List<ExcelChartAxis>();
        foreach (XmlNode node in this._chartNode.ParentNode.ChildNodes)
        {
            if (node.LocalName.EndsWith("Ax"))
            {
                ExcelChartAxis ax = new ExcelChartAxisStandard(this, this.NameSpaceManager, node, "c");
                l.Add(ax);
            }
        }

        this._axis = l.ToArray();

        XmlNodeList nl = this._chartNode.SelectNodes("c:axId", this.NameSpaceManager);
        foreach (XmlNode node in nl)
        {                
            string id = node.Attributes["val"].Value;
            int ix = Array.FindIndex(this._axis, x => x.Id == id);
            if(ix>=0)
            {
                if(this.XAxis==null)
                {
                    this.XAxis = (ExcelChartAxisStandard)this._axis[ix];
                }
                else
                {
                    this.YAxis = (ExcelChartAxisStandard)this._axis[ix];
                    break;
                }
            }
        }
    }
    internal virtual eChartType GetChartType(string name)
    {
        switch (name)
        {
            case "area3DChart":
                if (this.Grouping == eGrouping.Stacked)
                {
                    return eChartType.AreaStacked3D;
                }
                else if (this.Grouping == eGrouping.PercentStacked)
                {
                    return eChartType.AreaStacked1003D;
                }
                else
                {
                    return eChartType.Area3D;
                }
            case "areaChart":
                if (this.Grouping == eGrouping.Stacked)
                {
                    return eChartType.AreaStacked;
                }
                else if (this.Grouping == eGrouping.PercentStacked)
                {
                    return eChartType.AreaStacked100;
                }
                else
                {
                    return eChartType.Area;
                }
            case "doughnutChart":
                return eChartType.Doughnut;
            case "pie3DChart":
                return eChartType.Pie3D;
            case "pieChart":
                return eChartType.Pie;
            case "radarChart":
                return eChartType.Radar;
            case "scatterChart":
                return eChartType.XYScatter;
            case "surface3DChart":
            case "surfaceChart":
                return eChartType.Surface;
            case "stockChart":
                return eChartType.StockHLC;
            default:
                return 0;
        }
    }
    #region "Xml init Functions"
    private string ChartStartXml(eChartType type)
    {
        StringBuilder xml = new StringBuilder();
        int axID = 1;
        int xAxID = 2;
        int serAxID = this.HasThirdAxis() ? 3 : -1;

        xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        xml.AppendFormat("<c:chartSpace xmlns:c=\"{0}\" xmlns:a=\"{1}\" xmlns:r=\"{2}\">", ExcelPackage.schemaChart, ExcelPackage.schemaDrawings, ExcelPackage.schemaRelationships);
        xml.Append("<c:chart>");
        xml.AppendFormat("{0}{1}<c:plotArea><c:layout/>", this.AddPerspectiveXml(type), this.Add3DXml(type));

        string chartNodeText = this.GetChartNodeText();
        if(type==eChartType.StockVHLC || type==eChartType.StockVOHLC)
        {
            this.AppendStockChartXml(type, xml, chartNodeText);
        }
        else
        {
            xml.AppendFormat("<{0}>", chartNodeText);
            xml.Append(this.GetChartSerieStartXml(type, axID, xAxID, serAxID));
            xml.AppendFormat("</{0}>", chartNodeText);
        }

        //Axis
        if (!this.IsTypePieDoughnut())
        {
            if (this.IsTypeScatter() || this.IsTypeBubble())
            {
                xml.AppendFormat("<c:valAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/>{2}<c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/></c:valAx>", axID, xAxID, GetAxisShapeProperties());
            }
            else
            {
                xml.AppendFormat("<c:catAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/>{2}<c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx>", axID, xAxID, GetAxisShapeProperties());
            }
            xml.AppendFormat("<c:valAx><c:axId val=\"{1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:majorGridlines/><c:tickLblPos val=\"nextTo\"/>{2}<c:crossAx val=\"{0}\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx>", axID, xAxID, GetAxisShapeProperties());
            if (serAxID == 3) //Surface Chart
            {
                if (this.IsTypeSurface() || this.ChartType==eChartType.Line3D)
                {
                    xml.AppendFormat("<c:serAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/>{2}<c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/></c:serAx>", serAxID, xAxID, GetAxisShapeProperties());
                }
                else
                {
                    xml.AppendFormat("<c:valAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"r\"/><c:majorGridlines/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/>{2}<c:crossAx val=\"{1}\"/><c:crosses val=\"max\"/><c:crossBetween val=\"between\"/></c:valAx>", serAxID, axID, GetAxisShapeProperties());
                }
            }
        }

        xml.AppendFormat("</c:plotArea>" +      //<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>
                         AddLegend() +
                         "<c:plotVisOnly val=\"1\"/></c:chart>", axID, xAxID);

        xml.Append("<c:printSettings><c:headerFooter/><c:pageMargins b=\"0.75\" l=\"0.7\" r=\"0.7\" t=\"0.75\" header=\"0.3\" footer=\"0.3\"/><c:pageSetup/></c:printSettings></c:chartSpace>");
        return xml.ToString();
    }

    private void AppendStockChartXml(eChartType type, StringBuilder xml, string chartNodeText)
    {
        xml.Append("<c:barChart>");
        xml.Append(this.AddAxisId(1, 2, -1));
        xml.Append("</c:barChart>");
        xml.AppendFormat("<{0}>", chartNodeText);
        xml.Append(this.GetChartSerieStartXml(type, 1, 3, -1));
        xml.AppendFormat("</{0}>", chartNodeText);
    }

    private static object GetAxisShapeProperties()
    {
        return //"<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>" +
            "<c:txPr>" +
            "<a:bodyPr rot=\"-60000000\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\"/>" +
            "<a:lstStyle/>" +
            "<a:p><a:pPr><a:defRPr kern=\"1200\" sz=\"900\"/></a:pPr></a:p>" +
            "</c:txPr>";
    }

    private static string AddLegend()
    {
        return "<c:legend><c:legendPos val=\"r\"/><c:layout/><c:overlay val=\"0\" />" +
               //"<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>" +
               "<c:txPr><a:bodyPr anchorCtr=\"1\" anchor=\"ctr\" wrap=\"square\" vert=\"horz\" vertOverflow=\"ellipsis\" spcFirstLastPara=\"1\" rot=\"0\"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr/></a:p></c:txPr>" +
               "</c:legend>";
    }

    private string GetChartSerieStartXml(eChartType type, int axID, int xAxID, int serAxID)
    {
        StringBuilder xml = new StringBuilder();

        xml.Append(AddScatterType(type));
        xml.Append(AddRadarType(type));
        xml.Append(this.AddBarDir(type));
        xml.Append(this.AddGrouping());
        xml.Append(this.AddVaryColors());
        xml.Append(AddHasMarker(type));
        xml.Append(this.AddShape(type));
        xml.Append(AddFirstSliceAng(type));
        xml.Append(AddHoleSize(type));
        if (this.ChartType == eChartType.BarStacked100 || this.ChartType == eChartType.BarStacked || this.ChartType == eChartType.ColumnStacked || this.ChartType == eChartType.ColumnStacked100)
        {
            xml.Append("<c:overlap val=\"100\"/>");
        }
        if (this.IsTypeSurface())
        {
            xml.Append("<c:bandFmts/>");
        }
        xml.Append(this.AddAxisId(axID, xAxID, serAxID));

        return xml.ToString();
    }
    private string AddAxisId(int axID, int xAxID, int serAxID)
    {
        if (!this.IsTypePieDoughnut())
        {
            if (serAxID>0)
            {
                return string.Format("<c:axId val=\"{0}\"/><c:axId val=\"{1}\"/><c:axId val=\"{2}\"/>", axID, xAxID, serAxID);
            }
            else
            {
                return string.Format("<c:axId val=\"{0}\"/><c:axId val=\"{1}\"/>", axID, xAxID);
            }
        }
        else
        {
            return "";
        }
    }
    private string AddAxType()
    {
        switch (this.ChartType)
        {
            case eChartType.XYScatter:
            case eChartType.XYScatterLines:
            case eChartType.XYScatterLinesNoMarkers:
            case eChartType.XYScatterSmooth:
            case eChartType.XYScatterSmoothNoMarkers:
            case eChartType.Bubble:
            case eChartType.Bubble3DEffect:
                return "valAx";
            default:
                return "catAx";
        }
    }
    private static string AddScatterType(eChartType type)
    {
        if (type == eChartType.XYScatter ||
            type == eChartType.XYScatterLines ||
            type == eChartType.XYScatterLinesNoMarkers ||
            type == eChartType.XYScatterSmooth ||
            type == eChartType.XYScatterSmoothNoMarkers)
        {
            return "<c:scatterStyle val=\"\" />";
        }
        else
        {
            return "";
        }
    }
    private static string AddRadarType(eChartType type)
    {
        if (type == eChartType.Radar ||
            type == eChartType.RadarFilled ||
            type == eChartType.RadarMarkers)
        {
            return "<c:radarStyle val=\"\" />";
        }
        else
        {
            return "";
        }
    }
    private string AddGrouping()
    {
        //IsTypeClustered() || IsTypePercentStacked() || IsTypeStacked() || 
        if (this.IsTypeShape() || this.IsTypeLine())
        {
            return "<c:grouping val=\"standard\"/>";
        }
        else
        {
            return "";
        }
    }
    private static string AddHoleSize(eChartType type)
    {
        if (type == eChartType.Doughnut ||
            type == eChartType.DoughnutExploded)
        {
            return "<c:holeSize val=\"50\" />";
        }
        else
        {
            return "";
        }
    }
    private static string AddFirstSliceAng(eChartType type)
    {
        if (type == eChartType.Doughnut ||
            type == eChartType.DoughnutExploded)
        {
            return "<c:firstSliceAng val=\"0\" />";
        }
        else
        {
            return "";
        }
    }
    private string AddVaryColors()
    {
        if (this.IsTypeStock() || this.IsTypeSurface())
        {
            return "";
        }
        else
        {
            if (this.IsTypePieDoughnut())
            {
                return "<c:varyColors val=\"1\" />";
            }
            else
            {
                return "<c:varyColors val=\"0\" />";
            }
        }
    }
    private static string AddHasMarker(eChartType type)
    {
        if (type == eChartType.LineMarkers ||
            type == eChartType.LineMarkersStacked ||
            type == eChartType.LineMarkersStacked100 /*||
               type == eChartType.XYScatterLines ||
               type == eChartType.XYScatterSmooth*/)
        {
            return "<c:marker val=\"1\"/>";
        }
        else
        {
            return "";
        }
    }
    private string AddShape(eChartType type)
    {
        if (this.IsTypeShape())
        {
            return "<c:shape val=\"box\" />";
        }
        else
        {
            return "";
        }
    }
    private string AddBarDir(eChartType type)
    {
        if (this.IsTypeShape())
        {
            return "<c:barDir val=\"col\" />";
        }
        else
        {
            return "";
        }
    }
    private string AddPerspectiveXml(eChartType type)
    {
        //Add for 3D sharts
        if (this.IsType3D())
        {
            return "<c:view3D><c:perspective val=\"30\" /></c:view3D>";
        }
        else
        {
            return "";
        }
    }
    private string Add3DXml(eChartType type)
    {
        if (this.IsType3D())
        {
            return Add3DPart("floor") + Add3DPart("sideWall") + Add3DPart("backWall");
        }
        else
        {
            return "";
        }
    }

    private static string Add3DPart(string name)
    {
        return string.Format("<c:{0}><c:thickness val=\"0\"/></c:{0}>", name);  //<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/><a:sp3d/></c:spPr>
    }
    #endregion
    #endregion

    /// <summary>
    /// Get the name of the chart node
    /// </summary>
    /// <returns>The name</returns>
    protected string GetChartNodeText()
    {
        switch (this.ChartType)
        {
            case eChartType.Area3D:
            case eChartType.AreaStacked3D:
            case eChartType.AreaStacked1003D:
                return "c:area3DChart";
            case eChartType.Area:
            case eChartType.AreaStacked:
            case eChartType.AreaStacked100:
                return "c:areaChart";
            case eChartType.BarClustered:
            case eChartType.BarStacked:
            case eChartType.BarStacked100:
            case eChartType.ColumnClustered:
            case eChartType.ColumnStacked:
            case eChartType.ColumnStacked100:
                return "c:barChart";
            case eChartType.Column3D:
            case eChartType.BarClustered3D:
            case eChartType.BarStacked3D:
            case eChartType.BarStacked1003D:
            case eChartType.ColumnClustered3D:
            case eChartType.ColumnStacked3D:
            case eChartType.ColumnStacked1003D:
            case eChartType.ConeBarClustered:
            case eChartType.ConeBarStacked:
            case eChartType.ConeBarStacked100:
            case eChartType.ConeCol:
            case eChartType.ConeColClustered:
            case eChartType.ConeColStacked:
            case eChartType.ConeColStacked100:
            case eChartType.CylinderBarClustered:
            case eChartType.CylinderBarStacked:
            case eChartType.CylinderBarStacked100:
            case eChartType.CylinderCol:
            case eChartType.CylinderColClustered:
            case eChartType.CylinderColStacked:
            case eChartType.CylinderColStacked100:
            case eChartType.PyramidBarClustered:
            case eChartType.PyramidBarStacked:
            case eChartType.PyramidBarStacked100:
            case eChartType.PyramidCol:
            case eChartType.PyramidColClustered:
            case eChartType.PyramidColStacked:
            case eChartType.PyramidColStacked100:
                return "c:bar3DChart";
            case eChartType.Bubble:
            case eChartType.Bubble3DEffect:
                return "c:bubbleChart";
            case eChartType.Doughnut:
            case eChartType.DoughnutExploded:
                return "c:doughnutChart";
            case eChartType.Line:
            case eChartType.LineMarkers:
            case eChartType.LineMarkersStacked:
            case eChartType.LineMarkersStacked100:
            case eChartType.LineStacked:
            case eChartType.LineStacked100:
                return "c:lineChart";
            case eChartType.Line3D:
                return "c:line3DChart";
            case eChartType.Pie:
            case eChartType.PieExploded:
                return "c:pieChart";
            case eChartType.BarOfPie:
            case eChartType.PieOfPie:
                return "c:ofPieChart";
            case eChartType.Pie3D:
            case eChartType.PieExploded3D:
                return "c:pie3DChart";
            case eChartType.Radar:
            case eChartType.RadarFilled:
            case eChartType.RadarMarkers:
                return "c:radarChart";
            case eChartType.XYScatter:
            case eChartType.XYScatterLines:
            case eChartType.XYScatterLinesNoMarkers:
            case eChartType.XYScatterSmooth:
            case eChartType.XYScatterSmoothNoMarkers:
                return "c:scatterChart";
            case eChartType.Surface:
            case eChartType.SurfaceWireframe:
                return "c:surface3DChart";
            case eChartType.SurfaceTopView:
            case eChartType.SurfaceTopViewWireframe:
                return "c:surfaceChart";
            case eChartType.StockHLC:
            case eChartType.StockOHLC:
            case eChartType.StockVHLC:
            case eChartType.StockVOHLC:
                return "c:stockChart";
            default:
                throw (new NotImplementedException("Chart type not implemented"));
        }
    }
    /// <summary>
    /// Add a secondary axis
    /// </summary>
    internal override void AddAxis()
    {
        XmlElement catAx = this.ChartXml.CreateElement(string.Format("c:{0}", this.AddAxType()), ExcelPackage.schemaChart);
        int axID;
        if (this._axis.Length == 0)
        {
            this._plotArea.TopNode.AppendChild(catAx);
            axID = 1;
        }
        else
        {
            this._axis[0].TopNode.ParentNode.InsertAfter(catAx, this._axis[this._axis.Length - 1].TopNode);
            axID = int.Parse(this._axis[0].Id) < int.Parse(this._axis[1].Id) ? int.Parse(this._axis[1].Id) + 1 : int.Parse(this._axis[0].Id) + 1;
        }


        XmlElement valAx = this.ChartXml.CreateElement("c:valAx", ExcelPackage.schemaChart);
        catAx.ParentNode.InsertAfter(valAx, catAx);

        if (this._axis.Length == 0)
        {
            catAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/>", axID, axID + 1);
            valAx.InnerXml = string.Format("<c:axId val=\"{1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"l\"/><c:majorGridlines/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{0}\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/>", axID, axID + 1);
        }
        else
        {
            catAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"1\" /><c:axPos val=\"b\"/><c:tickLblPos val=\"none\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/>", axID, axID + 1);
            valAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"r\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"max\"/><c:crossBetween val=\"between\"/>", axID + 1, axID);
        }

        if (this._axis.Length == 0)
        {
            this._axis = new ExcelChartAxis[2];
        }
        else
        {
            ExcelChartAxis[] newAxis = new ExcelChartAxis[this._axis.Length + 2];
            Array.Copy(this._axis, newAxis, this._axis.Length);
            this._axis = newAxis;
        }

        this._axis[this._axis.Length - 2] = new ExcelChartAxisStandard(this, this.NameSpaceManager, catAx, "c");
        this._axis[this._axis.Length - 1] = new ExcelChartAxisStandard(this, this.NameSpaceManager, valAx, "c");
        foreach (ExcelChart? chart in this._plotArea.ChartTypes)
        {
            chart._axis = this._axis;
        }
    }
    internal void RemoveSecondaryAxis()
    {
        throw (new NotImplementedException("Not yet implemented"));
    }
    /// <summary>
    /// Title of the chart
    /// </summary>
    public new ExcelChartTitleStandard Title
    {
        get
        {
            this._title ??= this.GetTitle();

            return (ExcelChartTitleStandard)this._title;
        }
    }
    internal override ExcelChartTitle GetTitle()
    {
        return new ExcelChartTitleStandard(this, this.NameSpaceManager, this.ChartXml.SelectSingleNode("c:chartSpace/c:chart", this.NameSpaceManager), "c");
    }
    /// <summary>
    /// True if the chart has a title
    /// </summary>
    public override bool HasTitle
    {
        get
        {
            return this.ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:title", this.NameSpaceManager) != null;
        }
    }
    /// <summary>
    /// If the chart has a legend
    /// </summary>
    public override bool HasLegend
    {
        get
        {
            return this.ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", this.NameSpaceManager) != null;
        }
    }
    /// <summary>
    /// Remove the title from the chart
    /// </summary>
    public override void DeleteTitle()
    {
        this._title = null;
        this._chartXmlHelper.DeleteNode("../../c:title");
    }
    /// <summary>
    /// The build-in chart styles. 
    /// </summary>
    public override eChartStyle Style
    {
        get
        {
            XmlNode node = this.ChartXml.SelectSingleNode("c:chartSpace/c:style/@val", this.NameSpaceManager);
            if (node == null)
            {
                return eChartStyle.None;
            }
            else
            {
                if (int.TryParse(node.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out int v))
                {
                    return (eChartStyle)v;
                }
                else
                {
                    return eChartStyle.None;
                }
            }
        }
        set
        {
            if (value == eChartStyle.None)
            {
                XmlElement element = this.ChartXml.SelectSingleNode("c:chartSpace/c:style", this.NameSpaceManager) as XmlElement;
                if (element != null)
                {
                    element.ParentNode.RemoveChild(element);
                }
            }
            else
            {
                if (!this._chartXmlHelper.ExistsNode("../../../c:style"))
                {
                    XmlElement element = this.ChartXml.CreateElement("c:style", ExcelPackage.schemaChart);
                    element.SetAttribute("val", ((int)value).ToString());
                    XmlElement parent = this.ChartXml.SelectSingleNode("c:chartSpace", this.NameSpaceManager) as XmlElement;
                    parent.InsertBefore(element, parent.SelectSingleNode("c:chart", this.NameSpaceManager));
                }
                else
                {
                    this._chartXmlHelper.SetXmlNodeString("../../../ c:style/@val", ((int)value).ToString(CultureInfo.InvariantCulture));
                }
            }
        }
    }
    const string _roundedCornersPath = "../../../c:roundedCorners/@val";
    /// <summary>
    /// Border rounded corners
    /// </summary>
    public override bool RoundedCorners
    {
        get
        {
            return this._chartXmlHelper.GetXmlNodeBool(_roundedCornersPath);
        }
        set
        {
            this._chartXmlHelper.SetXmlNodeBool(_roundedCornersPath, value);
        }
    }
    const string _plotVisibleOnlyPath = "../../c:plotVisOnly/@val";
    /// <summary>
    /// Show data in hidden rows and columns
    /// </summary>
    public override bool ShowHiddenData
    {
        get
        {
            //!!Inverted value!!
            return !this._chartXmlHelper.GetXmlNodeBool(_plotVisibleOnlyPath);
        }
        set
        {
            //!!Inverted value!!
            this._chartXmlHelper.SetXmlNodeBool(_plotVisibleOnlyPath, !value);
        }
    }
    const string _displayBlanksAsPath = "../../c:dispBlanksAs/@val";
    /// <summary>
    /// Specifies the possible ways to display blanks
    /// </summary>
    public override eDisplayBlanksAs DisplayBlanksAs
    {
        get
        {
            string v = this._chartXmlHelper.GetXmlNodeString(_displayBlanksAsPath);
            if (string.IsNullOrEmpty(v))
            {
                return eDisplayBlanksAs.Zero; //Issue 14715 Changed in Office 2010-?
            }
            else
            {
                return (eDisplayBlanksAs)Enum.Parse(typeof(eDisplayBlanksAs), v, true);
            }
        }
        set
        {
            this._chartXmlHelper.SetXmlNodeString(_displayBlanksAsPath, value.ToString().ToLower(CultureInfo.InvariantCulture));
        }
    }
    const string _showDLblsOverMax = "../../c:showDLblsOverMax/@val";
    /// <summary>
    /// Specifies data labels over the maximum of the chart shall be shown
    /// </summary>
    public override bool ShowDataLabelsOverMaximum
    {
        get
        {
            return this._chartXmlHelper.GetXmlNodeBool(_showDLblsOverMax, true);
        }
        set
        {
            this._chartXmlHelper.SetXmlNodeBool(_showDLblsOverMax, value, true);
        }
    }
    /// <summary>
    /// Remove all axis that are not used any more
    /// </summary>
    /// <param name="excelChartAxis"></param>
    private void CheckRemoveAxis(ExcelChartAxis excelChartAxis)
    {
        if (this.ExistsAxis(excelChartAxis))
        {
            //Remove the axis
            ExcelChartAxis[] newAxis = new ExcelChartAxis[this.Axis.Length - 1];
            int pos = 0;
            foreach (ExcelChartAxisStandard? ax in this.Axis)
            {
                if (ax != excelChartAxis)
                {
                    newAxis[pos] = ax;
                }
            }

            //Update all charttypes.
            foreach (ExcelChart chartType in this._plotArea.ChartTypes)
            {
                chartType._axis = newAxis;
            }
        }
    }
    private bool ExistsAxis(ExcelChartAxis excelChartAxis)
    {
        foreach (ExcelChart chartType in this._plotArea.ChartTypes)
        {
            if (chartType != this)
            {
                if (chartType.XAxis.AxisPosition == excelChartAxis.AxisPosition ||
                    chartType.YAxis.AxisPosition == excelChartAxis.AxisPosition)
                {
                    //The axis exists
                    return true;
                }
            }
        }
        return false;
    }
    /// <summary>
    /// Plotarea
    /// </summary>
    public override ExcelChartPlotArea PlotArea
    {
        get
        {
            return this._plotArea ??= new ExcelChartPlotArea(this.NameSpaceManager,
                                                             this.ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:plotArea", this.NameSpaceManager),
                                                             this,
                                                             "c");
        }
    }
    /// <summary>
    /// Legend
    /// </summary>
    public new ExcelChartLegend Legend
    {
        get
        {
            return this._legend ??= new ExcelChartLegend(this.NameSpaceManager,
                                                         this.ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", this.NameSpaceManager),
                                                         this,
                                                         "c");
        }

    }
    ExcelDrawingBorder _border = null;
    /// <summary>
    /// Border
    /// </summary>
    public override ExcelDrawingBorder Border
    {
        get
        {
            return this._border ??= new ExcelDrawingBorder(this,
                                                           this.NameSpaceManager,
                                                           this.ChartXml.SelectSingleNode("c:chartSpace", this.NameSpaceManager),
                                                           "c:spPr/a:ln",
                                                           this._chartXmlHelper.SchemaNodeOrder);
        }
    }
    ExcelDrawingFill _fill = null;
    /// <summary>
    /// Access to Fill properties
    /// </summary>
    public override ExcelDrawingFill Fill
    {
        get
        {
            return this._fill ??= new ExcelDrawingFill(this,
                                                       this.NameSpaceManager,
                                                       this.ChartXml.SelectSingleNode("c:chartSpace", this.NameSpaceManager),
                                                       "c:spPr",
                                                       this._chartXmlHelper.SchemaNodeOrder);
        }
    }
    ExcelDrawingEffectStyle _effect = null;
    /// <summary>
    /// Effects
    /// </summary>
    public override ExcelDrawingEffectStyle Effect
    {
        get
        {
            return this._effect ??= new ExcelDrawingEffectStyle(this,
                                                                this.NameSpaceManager,
                                                                this.ChartXml.SelectSingleNode("c:chartSpace", this.NameSpaceManager),
                                                                "c:spPr/a:effectLst",
                                                                this._chartXmlHelper.SchemaNodeOrder);
        }
    }
    ExcelDrawing3D _threeD = null;
    /// <summary>
    /// 3D properties
    /// </summary>
    public override ExcelDrawing3D ThreeD
    {
        get
        {
            return this._threeD ??= new ExcelDrawing3D(this.NameSpaceManager,
                                                       this.ChartXml.SelectSingleNode("c:chartSpace", this.NameSpaceManager),
                                                       "c:spPr",
                                                       this._chartXmlHelper.SchemaNodeOrder);
        }
    }
    ExcelTextFont _font = null;
    /// <summary>
    /// Access to font properties
    /// </summary>
    public override ExcelTextFont Font
    {
        get
        {
            return this._font ??= new ExcelTextFont(this,
                                                    this.NameSpaceManager,
                                                    this.ChartXml.SelectSingleNode("c:chartSpace", this.NameSpaceManager),
                                                    "c:txPr/a:p/a:pPr/a:defRPr",
                                                    this._chartXmlHelper.SchemaNodeOrder);
        }
    }
    ExcelTextBody _textBody = null;
    /// <summary>
    /// Access to text body properties
    /// </summary>
    public override ExcelTextBody TextBody
    {
        get
        {
            return this._textBody ??= new ExcelTextBody(this.NameSpaceManager,
                                                        this.ChartXml.SelectSingleNode("c:chartSpace", this.NameSpaceManager),
                                                        "c:txPr/a:bodyPr",
                                                        this._chartXmlHelper.SchemaNodeOrder);
        }
    }

    /// <summary>
    /// 3D-settings
    /// </summary>
    public override ExcelView3D View3D
    {
        get
        {
            if (this.IsType3D())
            {
                return new ExcelView3D(this.NameSpaceManager, this.ChartXml.SelectSingleNode("//c:view3D", this.NameSpaceManager));
            }
            else
            {
                throw (new Exception("Charttype does not support 3D"));
            }

        }
    }
    string _groupingPath = "c:grouping/@val";
    /// <summary>
    /// Specifies the kind of grouping for a column, line, or area chart
    /// </summary>
    public eGrouping Grouping
    {
        get
        {
            return GetGroupingEnum(this._chartXmlHelper.GetXmlNodeString(this._groupingPath));
        }
        internal set
        {
            this._chartXmlHelper.SetXmlNodeString(this._groupingPath, GetGroupingText(value));
        }
    }
    string _varyColorsPath = "c:varyColors/@val";
    /// <summary>
    /// If the chart has only one serie this varies the colors for each point.
    /// </summary>
    public override bool VaryColors
    {
        get
        {
            return this._chartXmlHelper.GetXmlNodeBool(this._varyColorsPath);
        }
        set
        {
            if (value)
            {
                this._chartXmlHelper.SetXmlNodeString(this._varyColorsPath, "1");
            }
            else
            {
                this._chartXmlHelper.SetXmlNodeString(this._varyColorsPath, "0");
            }
        }
    }

    #region "Grouping Enum Translation"
    private static string GetGroupingText(eGrouping grouping)
    {
        switch (grouping)
        {
            case eGrouping.Clustered:
                return "clustered";
            case eGrouping.Stacked:
                return "stacked";
            case eGrouping.PercentStacked:
                return "percentStacked";
            default:
                return "standard";

        }
    }
    private static eGrouping GetGroupingEnum(string grouping)
    {
        switch (grouping)
        {
            case "stacked":
                return eGrouping.Stacked;
            case "percentStacked":
                return eGrouping.PercentStacked;
            default: //"clustered":               
                return eGrouping.Clustered;
        }
    }
    #endregion

    internal static int Items
    {
        get
        {
            return 0;
        }
    }

    internal void SetPivotSource(ExcelPivotTable pivotTableSource)
    {
        this.PivotTableSource = pivotTableSource;
        XmlElement chart = this.ChartXml.SelectSingleNode("c:chartSpace/c:chart", this.NameSpaceManager) as XmlElement;

        XmlElement? pivotSource = this.ChartXml.CreateElement("pivotSource", ExcelPackage.schemaChart);
        chart.ParentNode.InsertBefore(pivotSource, chart);
        pivotSource.InnerXml = string.Format("<c:name>[]{0}!{1}</c:name><c:fmtId val=\"0\"/>", this.PivotTableSource.WorkSheet.Name, pivotTableSource.Name);

        XmlElement? fmts = this.ChartXml.CreateElement("pivotFmts", ExcelPackage.schemaChart);
        chart.PrependChild(fmts);
        fmts.InnerXml = "<c:pivotFmt><c:idx val=\"0\"/><c:marker><c:symbol val=\"none\"/></c:marker></c:pivotFmt>";

        this.Series.AddPivotSerie(pivotTableSource);
    }
    ExcelChartAxisStandard[] _axisStandard = null;
    public new ExcelChartAxisStandard[] Axis
    {
        get { return this._axisStandard ??= this._axis.Select(x => (ExcelChartAxisStandard)x).ToArray(); }
    }
    /// <summary>
    /// The X Axis
    /// </summary>
    public new ExcelChartAxisStandard XAxis
    {
        get
        {
            return (ExcelChartAxisStandard)base.XAxis;
        }
        internal set
        {
            base.XAxis = value;
        }
    }
    /// <summary>
    /// The Y Axis
    /// </summary>
    public new ExcelChartAxisStandard YAxis
    {
        get
        {
            return (ExcelChartAxisStandard)base.YAxis;
        }
        internal set
        {
            base.YAxis = value;
        }
    }
}