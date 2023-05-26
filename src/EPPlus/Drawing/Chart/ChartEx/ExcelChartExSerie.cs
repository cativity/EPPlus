/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB           Release EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Globalization;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// A chart serie
    /// </summary>
    public class ExcelChartExSerie : ExcelChartSerie
    {
        XmlNode _dataNode;
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chart">The chart</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
        internal ExcelChartExSerie(ExcelChartEx chart, XmlNamespaceManager ns, XmlNode node)
            : base(chart, ns, node, "cx")
        {
            this.SchemaNodeOrder = new string[] { "tx", "spPr", "valueColors", "valueColorPositions", "dataPt", "dataLabels", "dataId", "layoutPr", "axisId" };
            this._dataNode = node.SelectSingleNode($"../../../../cx:chartData/cx:data[@id={this.DataId}]", ns);
            if((chart.ChartType == eChartType.BoxWhisker ||
                chart.ChartType == eChartType.Histogram ||
                chart.ChartType == eChartType.Pareto ||
                chart.ChartType == eChartType.Waterfall ||
                chart.ChartType == eChartType.Pareto) && chart.Series.Count==0)
            {
                if(chart._chartXmlHelper.ExistsNode("cx:plotArea/cx:axis")==false)
                {
                    this.AddAxis();
                }
                chart.LoadAxis();
            }
        }

        private void AddAxis()
        {
            XmlElement? axis0=(XmlElement)this._chart._chartXmlHelper.CreateNode("cx:plotArea/cx:axis");
            axis0.SetAttribute("id", "0");
            XmlElement? axis1 = (XmlElement)this._chart._chartXmlHelper.CreateNode("cx:plotArea/cx:axis", false, true);
            axis1.SetAttribute("id", "1");

            switch(this._chart.ChartType)
            {
                case eChartType.BoxWhisker:
                    axis0.InnerXml = "<cx:catScaling gapWidth=\"1\"/><cx:tickLabels/>";
                    axis1.InnerXml = "<cx:valScaling/><cx:majorGridlines/><cx:tickLabels/>";
                    break;
                case eChartType.Waterfall:
                    axis0.InnerXml = "<cx:catScaling/><cx:tickLabels/>";
                    axis1.InnerXml = "<cx:valScaling/><cx:tickLabels/>";
                    break;
                case eChartType.Funnel:
                    axis1.InnerXml = "<cx:catScaling gapWidth=\"0.06\"/><cx:tickLabels/>";
                    break;
                case eChartType.Histogram:
                case eChartType.Pareto:
                    axis0.InnerXml = "<cx:catScaling gapWidth=\"0\"/><cx:tickLabels/>";
                    axis1.InnerXml = "<cx:valScaling/><cx:majorGridlines/><cx:tickLabels/>";
                    if(this._chart.ChartType== eChartType.Pareto)
                    {
                        XmlElement? axis2 = (XmlElement)this._chart._chartXmlHelper.CreateNode("cx:plotArea/cx:axis", false, true);
                        axis2.SetAttribute("id", "2");
                        axis2.InnerXml = "<cx:valScaling min=\"0\" max=\"1\"/><cx:units unit=\"percentage\"/><cx:tickLabels/>";
                    }
                    break;
            }
        }
        internal int DataId
        {
            get
            {
                return this.GetXmlNodeInt("cx:dataId/@val");
            }
        }
        ExcelChartExDataCollection _dataDimensions = null;
        /// <summary>
        /// The dimensions of the serie
        /// </summary>
        public ExcelChartExDataCollection DataDimensions
        {
            get
            {
                if (this._dataDimensions == null)
                {
                    this._dataDimensions = new ExcelChartExDataCollection(this, this.NameSpaceManager, this._dataNode);
                }
                return this._dataDimensions;
            }
        }
        const string headerAddressPath = "c:tx/c:strRef/c:f";
        /// <summary>
        /// Header address for the serie.
        /// </summary>
        public override ExcelAddressBase HeaderAddress
        {
            get
            {
                string? f = this.GetXmlNodeString("cx:tx/cx:txData/cx:f");
                if (ExcelCellBase.IsValidAddress(f))
                {
                    return new ExcelAddressBase(f);
                }
                else
                {
                    if (this._chart.WorkSheet.Workbook.Names.ContainsKey(f))
                    {
                        return this._chart.WorkSheet.Workbook.Names[f];
                    }
                    else if (this._chart.WorkSheet.Names.ContainsKey(f))
                    {
                        return this._chart.WorkSheet.Names[f];
                    }
                    return null;
                }
            }
            set
            {
                this.SetXmlNodeString("cx:tx/cx:txData/cx:f", value.FullAddress);
            }
        }
        /// <summary>
        /// The header text for the serie.
        /// </summary>
        public override string Header
        {
            get
            {
                return this.GetXmlNodeString("cx:tx/cx:txData/cx:v");
            }
            set
            {
                this.SetXmlNodeString("cx:tx/cx:txData/cx:v", value);
            }
        }
        XmlHelper _catSerieHelper = null;
        XmlHelper _valSerieHelper = null;
        /// <summary>
        /// Set this to a valid address or the drawing will be invalid.
        /// </summary>
        public override string Series
        {
            get
            {
                XmlHelper? helper = this.GetSerieHelper();
                return helper.GetXmlNodeString("cx:f");
            }
            set
            {
                XmlHelper? helper = this.GetSerieHelper();
                helper.SetXmlNodeString("cx:f", this.ToFullAddress(value));
            }
        }

        /// <summary>
        /// Set an address for the horizontal labels
        /// </summary>
        public override string XSeries
        {
            get
            {
                XmlHelper? helper = this.GetXSerieHelper(false);
                if(helper==null)
                {
                    return "";
                }
                else
                {
                    return helper.GetXmlNodeString("cx:f");
                }
            }
            set
            {
                XmlHelper? helper = this.GetXSerieHelper(true);
                helper.SetXmlNodeString("cx:f", this.ToFullAddress(value));
            }
        }
        private XmlHelper GetSerieHelper()
        {
            if (this._valSerieHelper == null)
            {
                if (this._dataNode.ChildNodes.Count == 1)
                {
                    this._valSerieHelper = XmlHelperFactory.Create(this.NameSpaceManager, this._dataNode.FirstChild);
                }
                else if (this._dataNode.ChildNodes.Count > 1)
                {
                    this._valSerieHelper = XmlHelperFactory.Create(this.NameSpaceManager, this._dataNode.ChildNodes[1]); 
                }
            }
            return this._valSerieHelper;
        }

        private XmlHelper GetXSerieHelper(bool create)
        {
            if (this._catSerieHelper == null)
            {
                if (this._dataNode.ChildNodes.Count == 1)
                {
                    if (create)
                    {
                        XmlElement? node = this._dataNode.OwnerDocument.CreateElement("cx", "strDim", ExcelPackage.schemaChartExMain);
                        this._dataNode.InsertBefore(node, this._dataNode.FirstChild);
                        this._catSerieHelper = XmlHelperFactory.Create(this.NameSpaceManager, node);
                    }
                    else
                    {
                        return null;
                    }
                }
                else if (this._dataNode.ChildNodes.Count > 1)
                {
                    this._catSerieHelper = XmlHelperFactory.Create(this.NameSpaceManager, this._dataNode.ChildNodes[0]); 
                }
            }
            return this._catSerieHelper;
        }

        ExcelChartExSerieDataLabel _dataLabels = null;
        /// <summary>
        /// Data label properties
        /// </summary>
        public ExcelChartExSerieDataLabel DataLabel
        {
            get
            {
                if (this._dataLabels == null)
                {
                    this._dataLabels = new ExcelChartExSerieDataLabel(this, this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder);
                }
                return this._dataLabels;
            }
        }
        ExcelChartExDataPointCollection _dataPoints = null;
        /// <summary>
        /// A collection of individual data points
        /// </summary>
        public ExcelChartExDataPointCollection DataPoints
        {
            get
            {
                if(this._dataPoints==null)
                {
                    this._dataPoints = new ExcelChartExDataPointCollection(this, this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder);
                }
                return this._dataPoints;
            }
        }
        /// <summary>
        /// If the serie is hidden
        /// </summary>
        public bool Hidden
        {
            get
            {
                return this.GetXmlNodeBool("@hidden", false);
            }
            set
            {
                this.SetXmlNodeBool("@hidden", value, false);
            }
        }

        /// <summary>
        /// If the chart has datalabel
        /// </summary>
        public bool HasDataLabel
        {
            get
            {
                return this.TopNode.SelectSingleNode("c:dataLabels", this.NameSpaceManager) != null;
            }
        }

        /// <summary>
        /// Number of items. Will always return 0, as no item data is stored.
        /// </summary>
        public override int NumberOfItems => 0;

        /// <summary>
        /// Trendline do not apply to extended charts.
        /// </summary>
        public override ExcelChartTrendlineCollection TrendLines => new ExcelChartTrendlineCollection(null);

        internal override void SetID(string id)
        {
            throw new NotImplementedException();
        }

        internal static XmlElement CreateSeriesAndDataElement(ExcelChartEx chart, bool hasCatSerie)
        {
            XmlElement ser = CreateSeriesElement(chart, chart.ChartType, chart.Series.Count);
            ser.InnerXml = $"<cx:dataId val=\"{chart.Series.Count}\"/><cx:layoutPr/>{AddAxisReferense(chart)}";
            SetLayoutProperties(chart, ser);

            chart._chartXmlHelper.CreateNode("../cx:chartData", true);
            XmlElement? dataElement = (XmlElement)chart._chartXmlHelper.CreateNode("../cx:chartData/cx:data", false, true);
            dataElement.SetAttribute("id", chart.Series.Count.ToString());
            string? innerXml="";
            if (hasCatSerie == true)
            {
                innerXml += $"<cx:strDim type=\"cat\"><cx:f></cx:f><cx:nf></cx:nf></cx:strDim>";
            }
            innerXml += $"<cx:numDim type=\"{GetNumType(chart.ChartType)}\"><cx:f></cx:f><cx:nf></cx:nf></cx:numDim>";
            dataElement.InnerXml = innerXml;
            return ser;
        }

        internal static XmlElement CreateSeriesElement(ExcelChartEx chart, eChartType type, int index, XmlNode referenceNode = null, bool isPareto=false)
        {
            XmlNode? plotareaNode = chart._chartXmlHelper.CreateNode("cx:plotArea/cx:plotAreaRegion");
            XmlElement? ser = plotareaNode.OwnerDocument.CreateElement("cx", "series", ExcelPackage.schemaChartExMain);
            XmlNodeList node = plotareaNode.SelectNodes("cx:series", chart.NameSpaceManager);

            if(node.Count > 0)
            {
                plotareaNode.InsertAfter(ser, referenceNode ?? node[node.Count - 1]);
            }
            else
            {
                XmlHelper? f = XmlHelperFactory.Create(chart.NameSpaceManager, plotareaNode);
                InserAfter(plotareaNode, "cx:plotSurface", ser);
            }
            ser.SetAttribute("formatIdx", index.ToString());
            ser.SetAttribute("uniqueId", "{" + Guid.NewGuid().ToString() + "}");
            ser.SetAttribute("layoutId", GetLayoutId(type, isPareto));
            return ser;
        }

        private static object AddAxisReferense(ExcelChartEx chart)
        {
            if(chart.ChartType==eChartType.Pareto)
            {
                return "<cx:axisId val=\"1\"/>";
            }
            else
            {
                return "";
            }            
        }

        private static string GetLayoutId(eChartType chartType, bool isPareto)
        {
            if (isPareto)
            {
                return "paretoLine";
            }

            switch(chartType)
            {
                case eChartType.Histogram:
                case eChartType.Pareto:
                    return "clusteredColumn";
                default:
                    return chartType.ToEnumString();
            }            
        }

        private static void SetLayoutProperties(ExcelChartEx chart, XmlElement ser)
        {
            XmlNode? layoutPr = ser.SelectSingleNode("cx:layoutPr", chart.NameSpaceManager);
            switch (chart.ChartType)
            {
                case eChartType.BoxWhisker:
                    layoutPr.InnerXml = "<cx:parentLabelLayout val=\"banner\"/><cx:visibility outliers=\"1\" nonoutliers=\"0\" meanMarker=\"1\" meanLine=\"0\"/><cx:statistics quartileMethod=\"exclusive\"/>";
                    break;
                case eChartType.Histogram:
                case eChartType.Pareto:
                    layoutPr.InnerXml = "<cx:binning intervalClosed=\"r\"/>";
                    break;
                case eChartType.RegionMap:
                    CultureInfo? ci = CultureInfo.CurrentCulture;
                    layoutPr.InnerXml = $"<cx:geography attribution = \"Powered by Bing\" cultureRegion = \"{ci.TwoLetterISOLanguageName}\" cultureLanguage = \"{ci.Name}\" ><cx:geoCache provider=\"{{E9337A44-BEBE-4D9F-B70C-5C5E7DAFC167}}\"><cx:binary/></cx:geoCache></cx:geography>";
                    break;
                case eChartType.Waterfall:
                    layoutPr.InnerXml = "<cx:visibility connectorLines=\"0\" />";
                    break;
            }
        }

        private static string GetNumType(eChartType chartType)
        {
            switch (chartType)
            {
                case eChartType.Sunburst:
                case eChartType.Treemap:
                    return "size";
                case eChartType.RegionMap:
                    return "colorVal";
                default:
                    return "val";
            }
        }
    }
}
