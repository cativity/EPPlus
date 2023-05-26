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
using System.Text;
using System.Xml;
using OfficeOpenXml.Table.PivotTable;
using System.Globalization;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Represents a Bar Chart
    /// </summary>
    public sealed class ExcelBarChart : ExcelChartStandard, IDrawingDataLabel
    {
        #region "Constructors"
        internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml, ExcelGroupShape parent = null) :
            base(drawings, node, type, topChart, PivotTableSource, chartXml, parent)
        {
            this.SetChartNodeText("");
            if(type.HasValue)
            {
                this.SetTypeProperties(drawings, type.Value);
            }
        }

        internal ExcelBarChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
           base(drawings, node, uriChart, part, chartXml, chartNode, parent)
        {
            this.SetChartNodeText(chartNode.Name);
        }

        internal ExcelBarChart(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent = null) : 
            base(topChart, chartNode, parent)
        {
            this.SetChartNodeText(chartNode.Name);
        }
        #endregion
        internal override void InitSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list = null)
        {
            base.InitSeries(chart, ns, node, isPivot, list);
            this.Series.Init(chart, ns, node, isPivot, base.Series._list);
        }

        #region "Private functions"
        //string _chartTopPath="c:chartSpace/c:chart/c:plotArea/{0}";
        private void SetChartNodeText(string chartNodeText)
        {
            if(string.IsNullOrEmpty(chartNodeText))
            {
                chartNodeText = this.GetChartNodeText();
            }
            //_chartTopPath = string.Format(_chartTopPath, chartNodeText);
            //_directionPath = string.Format(_directionPath, _chartTopPath);
            //_shapePath = string.Format(_shapePath, _chartTopPath);
        }
        private void SetTypeProperties(ExcelDrawings drawings, eChartType type)
        {
            /******* Bar direction *******/
            if (type == eChartType.BarClustered ||
                type == eChartType.BarStacked ||
                type == eChartType.BarStacked100 ||
                type == eChartType.BarClustered3D ||
                type == eChartType.BarStacked3D ||
                type == eChartType.BarStacked1003D ||
                type == eChartType.ConeBarClustered ||
                type == eChartType.ConeBarStacked ||
                type == eChartType.ConeBarStacked100 ||
                type == eChartType.CylinderBarClustered ||
                type == eChartType.CylinderBarStacked ||
                type == eChartType.CylinderBarStacked100 ||
                type == eChartType.PyramidBarClustered ||
                type == eChartType.PyramidBarStacked ||
                type == eChartType.PyramidBarStacked100)
            {
                this.Direction = eDirection.Bar;
            }
            else if (
                type == eChartType.ColumnClustered ||
                type == eChartType.ColumnStacked ||
                type == eChartType.ColumnStacked100 ||
                type == eChartType.Column3D ||
                type == eChartType.ColumnClustered3D ||
                type == eChartType.ColumnStacked3D ||
                type == eChartType.ColumnStacked1003D ||
                type == eChartType.ConeCol ||
                type == eChartType.ConeColClustered ||
                type == eChartType.ConeColStacked ||
                type == eChartType.ConeColStacked100 ||
                type == eChartType.CylinderCol ||
                type == eChartType.CylinderColClustered ||
                type == eChartType.CylinderColStacked ||
                type == eChartType.CylinderColStacked100 ||
                type == eChartType.PyramidCol ||
                type == eChartType.PyramidColClustered ||
                type == eChartType.PyramidColStacked ||
                type == eChartType.PyramidColStacked100)
            {
                this.Direction = eDirection.Column;
            }

            /****** Shape ******/
            if (/*type == eChartType.ColumnClustered ||
                type == eChartType.ColumnStacked ||
                type == eChartType.ColumnStacked100 ||*/
                type == eChartType.Column3D ||
                type == eChartType.ColumnClustered3D ||
                type == eChartType.ColumnStacked3D ||
                type == eChartType.ColumnStacked1003D ||
                /*type == eChartType.BarClustered ||
                type == eChartType.BarStacked ||
                type == eChartType.BarStacked100 ||*/
                type == eChartType.BarClustered3D ||
                type == eChartType.BarStacked3D ||
                type == eChartType.BarStacked1003D)
            {
                this.Shape = eShape.Box;
            }
            else if (
                type == eChartType.CylinderBarClustered ||
                type == eChartType.CylinderBarStacked ||
                type == eChartType.CylinderBarStacked100 ||
                type == eChartType.CylinderCol ||
                type == eChartType.CylinderColClustered ||
                type == eChartType.CylinderColStacked ||
                type == eChartType.CylinderColStacked100)
            {
                this.Shape = eShape.Cylinder;
            }
            else if (
                type == eChartType.ConeBarClustered ||
                type == eChartType.ConeBarStacked ||
                type == eChartType.ConeBarStacked100 ||
                type == eChartType.ConeCol ||
                type == eChartType.ConeColClustered ||
                type == eChartType.ConeColStacked ||
                type == eChartType.ConeColStacked100)
            {
                this.Shape = eShape.Cone;
            }
            else if (
                type == eChartType.PyramidBarClustered ||
                type == eChartType.PyramidBarStacked ||
                type == eChartType.PyramidBarStacked100 ||
                type == eChartType.PyramidCol ||
                type == eChartType.PyramidColClustered ||
                type == eChartType.PyramidColStacked ||
                type == eChartType.PyramidColStacked100)
            {
                this.Shape = eShape.Pyramid;
            }
        }
        #endregion
        #region "Properties"
        string _directionPath = "c:barDir/@val";
        /// <summary>
        /// Direction, Bar or columns
        /// </summary>
        public eDirection Direction
        {
            get
            {
                return GetDirectionEnum(this._chartXmlHelper.GetXmlNodeString(this._directionPath));
            }
            internal set
            {
                this._chartXmlHelper.SetXmlNodeString(this._directionPath, GetDirectionText(value));
            }
        }
        string _shapePath = "c:shape/@val";
        /// <summary>
        /// The shape of the bar/columns
        /// </summary>
        public eShape Shape
        {
            get
            {
                return GetShapeEnum(this._chartXmlHelper.GetXmlNodeString(this._shapePath));
            }
            internal set
            {
                this._chartXmlHelper.SetXmlNodeString(this._shapePath, GetShapeText(value));
            }
        }
        ExcelChartDataLabel _DataLabel = null;
        /// <summary>
        /// Access to datalabel properties
        /// </summary>
        public ExcelChartDataLabel DataLabel
        {
            get
            {
                if (this._DataLabel == null)
                {
                    this._DataLabel = new ExcelChartDataLabelStandard(this, this.NameSpaceManager, this.ChartNode, "dLbls", this._chartXmlHelper.SchemaNodeOrder);
                }
                return this._DataLabel;
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
        string _gapWidthPath = "c:gapWidth/@val";
        /// <summary>
        /// The size of the gap between two adjacent bars/columns
        /// </summary>
        public int GapWidth
        {
            get
            {
                return this._chartXmlHelper.GetXmlNodeInt(this._gapWidthPath);
            }
            set
            {
                this._chartXmlHelper.SetXmlNodeString(this._gapWidthPath, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        string _overlapPath = "c:overlap/@val";
        /// <summary>
        /// Specifies how much bars and columns shall overlap
        /// </summary>
        public int Overlap
        {
            get
            {
                return this._chartXmlHelper.GetXmlNodeInt(this._overlapPath);
            }
            set
            {
                this._chartXmlHelper.SetXmlNodeString(this._overlapPath, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        #endregion
        #region "Direction Enum Traslation"
        private static string GetDirectionText(eDirection direction)
        {
            switch (direction)
            {
                case eDirection.Bar:
                    return "bar";
                default:
                    return "col";
            }
        }

        private static eDirection GetDirectionEnum(string direction)
        {
            switch (direction)
            {
                case "bar":
                    return eDirection.Bar;
                default:
                    return eDirection.Column;
            }
        }
        #endregion
        #region "Shape Enum Translation"
        private static string GetShapeText(eShape Shape)
        {
            switch (Shape)
            {
                case eShape.Box:
                    return "box";
                case eShape.Cone:
                    return "cone";
                case eShape.ConeToMax:
                    return "coneToMax";
                case eShape.Cylinder:
                    return "cylinder";
                case eShape.Pyramid:
                    return "pyramid";
                case eShape.PyramidToMax:
                    return "pyramidToMax";
                default:
                    return "box";
            }
        }

        private static eShape GetShapeEnum(string text)
        {
            switch (text)
            {
                case "box":
                    return eShape.Box;
                case "cone":
                    return eShape.Cone;
                case "coneToMax":
                    return eShape.ConeToMax;
                case "cylinder":
                    return eShape.Cylinder;
                case "pyramid":
                    return eShape.Pyramid;
                case "pyramidToMax":
                    return eShape.PyramidToMax;
                default:
                    return eShape.Box;
            }
        }
        #endregion
        internal override eChartType GetChartType(string name)
        {
            if (name == "barChart")
            {
                if (this.Direction == eDirection.Bar)
                {
                    if (this.Grouping == eGrouping.Stacked)
                    {
                        return eChartType.BarStacked;
                    }
                    else if (this.Grouping == eGrouping.PercentStacked)
                    {
                        return eChartType.BarStacked100;
                    }
                    else
                    {
                        return eChartType.BarClustered;
                    }
                }
                else
                {
                    if (this.Grouping == eGrouping.Stacked)
                    {
                        return eChartType.ColumnStacked;
                    }
                    else if (this.Grouping == eGrouping.PercentStacked)
                    {
                        return eChartType.ColumnStacked100;
                    }
                    else
                    {
                        return eChartType.ColumnClustered;
                    }
                }
            }
            if (name == "bar3DChart")
            {
                #region "Bar Shape"
                if (this.Shape==eShape.Box)
                {
                    if (this.Direction == eDirection.Bar)
                    {
                        if (this.Grouping == eGrouping.Stacked)
                        {
                            return eChartType.BarStacked3D;
                        }
                        else if (this.Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.BarStacked1003D;
                        }
                        else
                        {
                            return eChartType.BarClustered3D;
                        }
                    }
                    else
                    {
                        if (this.Grouping == eGrouping.Stacked)
                        {
                            return eChartType.ColumnStacked3D;
                        }
                        else if (this.Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.ColumnStacked1003D;
                        }
                        else
                        {
                            return eChartType.ColumnClustered3D;
                        }
                    }
                }
                #endregion
                #region "Cone Shape"
                if (this.Shape == eShape.Cone || this.Shape == eShape.ConeToMax)
                {
                    if (this.Direction == eDirection.Bar)
                    {
                        if (this.Grouping == eGrouping.Stacked)
                        {
                            return eChartType.ConeBarStacked;
                        }
                        else if (this.Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.ConeBarStacked100;
                        }
                        else if (this.Grouping == eGrouping.Clustered)
                        {
                            return eChartType.ConeBarClustered;
                        }
                    }
                    else
                    {
                        if (this.Grouping == eGrouping.Stacked)
                        {
                            return eChartType.ConeColStacked;
                        }
                        else if (this.Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.ConeColStacked100;
                        }
                        else if (this.Grouping == eGrouping.Clustered)
                        {
                            return eChartType.ConeColClustered;
                        }
                        else
                        {
                            return eChartType.ConeCol;
                        }
                    }
                }
                #endregion
                #region "Cylinder Shape"
                if (this.Shape == eShape.Cylinder)
                {
                    if (this.Direction == eDirection.Bar)
                    {
                        if (this.Grouping == eGrouping.Stacked)
                        {
                            return eChartType.CylinderBarStacked;
                        }
                        else if (this.Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.CylinderBarStacked100;
                        }
                        else if (this.Grouping == eGrouping.Clustered)
                        {
                            return eChartType.CylinderBarClustered;
                        }
                    }
                    else
                    {
                        if (this.Grouping == eGrouping.Stacked)
                        {
                            return eChartType.CylinderColStacked;
                        }
                        else if (this.Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.CylinderColStacked100;
                        }
                        else if (this.Grouping == eGrouping.Clustered)
                        {
                            return eChartType.CylinderColClustered;
                        }
                        else
                        {
                            return eChartType.CylinderCol;
                        }
                    }
                }
                #endregion
                #region "Pyramide Shape"
                if (this.Shape == eShape.Pyramid || this.Shape == eShape.PyramidToMax)
                {
                    if (this.Direction == eDirection.Bar)
                    {
                        if (this.Grouping == eGrouping.Stacked)
                        {
                            return eChartType.PyramidBarStacked;
                        }
                        else if (this.Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.PyramidBarStacked100;
                        }
                        else if (this.Grouping == eGrouping.Clustered)
                        {
                            return eChartType.PyramidBarClustered;
                        }
                    }
                    else
                    {
                        if (this.Grouping == eGrouping.Stacked)
                        {
                            return eChartType.PyramidColStacked;
                        }
                        else if (this.Grouping == eGrouping.PercentStacked)
                        {
                            return eChartType.PyramidColStacked100;
                        }
                        else if (this.Grouping == eGrouping.Clustered)
                        {
                            return eChartType.PyramidColClustered;
                        }
                        else
                        {
                            return eChartType.PyramidCol;
                        }
                    }
                }
                #endregion
            }
            return base.GetChartType(name);
        }
        /// <summary>
        /// Series for a bar chart
        /// </summary>
        public new ExcelChartSeries<ExcelBarChartSerie> Series { get; } = new ExcelChartSeries<ExcelBarChartSerie>();
    }
}
