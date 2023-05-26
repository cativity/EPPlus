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
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Provides access to scatter chart specific properties
    /// </summary>
    public sealed class ExcelScatterChart : ExcelChartStandard, IDrawingDataLabel
    {
        internal ExcelScatterChart(ExcelDrawings drawings, XmlNode node, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml, ExcelGroupShape parent=null) :
            base(drawings, node, type, topChart, PivotTableSource, chartXml, parent)
        {
            this.SetTypeProperties();
        }

        internal ExcelScatterChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent) :
            base(drawings, node, uriChart, part, chartXml, chartNode, parent)
        {
            this.SetTypeProperties();
        }

        internal ExcelScatterChart(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent) :
            base(topChart, chartNode, parent)
        {
            this.SetTypeProperties();
        }
        internal override void InitSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list = null)
        {
            base.InitSeries(chart, ns, node, isPivot, list);
            this.Series.Init(chart, ns, node, isPivot, base.Series._list);
        }
        private void SetTypeProperties()
        {
           /***** ScatterStyle *****/
           if(this.ChartType == eChartType.XYScatter || this.ChartType == eChartType.XYScatterLines || this.ChartType == eChartType.XYScatterLinesNoMarkers)
           {
               this.ScatterStyle = eScatterStyle.LineMarker;
          }
           else if (this.ChartType == eChartType.XYScatterSmooth || this.ChartType == eChartType.XYScatterSmoothNoMarkers) 
           {
               this.ScatterStyle = eScatterStyle.SmoothMarker;
           }
        }
        #region "Grouping Enum Translation"
        string _scatterTypePath = "c:scatterStyle/@val";
        private static eScatterStyle GetScatterEnum(string text)
        {
            switch (text)
            {
                case "smoothMarker":
                    return eScatterStyle.SmoothMarker;
                default:
                    return eScatterStyle.LineMarker;
            }
        }

        private static string GetScatterText(eScatterStyle shatterStyle)
        {
            switch (shatterStyle)
            {
                case eScatterStyle.SmoothMarker:
                    return "smoothMarker";
                default:
                    return "lineMarker";
            }
        }
        #endregion
        /// <summary>
        /// If the scatter has LineMarkers or SmoothMarkers
        /// </summary>
        public eScatterStyle ScatterStyle
        {
            get
            {
                return GetScatterEnum(this._chartXmlHelper.GetXmlNodeString(this._scatterTypePath));
            }
            internal set
            {
                this._chartXmlHelper.CreateNode(this._scatterTypePath, true);
                this._chartXmlHelper.SetXmlNodeString(this._scatterTypePath, GetScatterText(value));
            }
        }
        string MARKER_PATH = "c:marker/@val";
        /// <summary>
        /// If the series has markers
        /// </summary>
        public bool Marker
        {
            get
            {
                return this.GetXmlNodeBool(this.MARKER_PATH, false);
            }
            set
            {
                this.SetXmlNodeBool(this.MARKER_PATH, value, false);
            }
        }
        ExcelChartDataLabel _dataLabel = null;
        /// <summary>
        /// Access to datalabel properties
        /// </summary>
        public ExcelChartDataLabel DataLabel
        {
            get
            {
                return this._dataLabel ??= new ExcelChartDataLabelStandard(this, this.NameSpaceManager, this.ChartNode, "dLbls", this._chartXmlHelper.SchemaNodeOrder);
            }
        }
        /// <summary>
        /// If the chart has datalabel
        /// </summary>
        public bool HasDataLabel
        {
            get
            {
                return this.ChartNode.SelectSingleNode("c:dLbls", this.NameSpaceManager) != null;
            }
        }
        internal override eChartType GetChartType(string name)
        {
            if (name == "scatterChart")
            {
                if (this.ScatterStyle==eScatterStyle.LineMarker)
                {
                    if (this.Series.Count > 0 && this.Series[0].Marker.Style == eMarkerStyle.None)
                    {
                        return eChartType.XYScatterLinesNoMarkers;
                    }
                    else
                    {
                        if(this.ExistsNode("c:ser/c:spPr/a:ln/noFill"))
                        {
                            return eChartType.XYScatter;
                        }
                        else
                        {
                            return eChartType.XYScatterLines;
                        }
                    }
                }
                else if (this.ScatterStyle == eScatterStyle.SmoothMarker)
                {
                    if (this.Series.Count > 0 && this.Series[0].Marker.Style == eMarkerStyle.None)
                    {
                        return eChartType.XYScatterSmoothNoMarkers;
                    }
                    else
                    {
                        return eChartType.XYScatterSmooth;
                    }
                }
            }
            return base.GetChartType(name);
        }
        /// <summary>
        /// A collection of series for a Scatter Chart
        /// </summary>
        public new ExcelChartSeries<ExcelScatterChartSerie> Series { get; } = new ExcelChartSeries<ExcelScatterChartSerie>();
    }
}
