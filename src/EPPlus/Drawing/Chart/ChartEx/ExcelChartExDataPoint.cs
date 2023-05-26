﻿/*************************************************************************************************
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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Style.ThreeD;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// An individual data point
    /// </summary>
    public class ExcelChartExDataPoint : XmlHelper, IDrawingStyleBase
    {
        ExcelChartExSerie _serie;
        internal ExcelChartExDataPoint(ExcelChartExSerie serie, XmlNamespaceManager ns, XmlNode topNode, string[] schemaNodeOrder) : base(ns, topNode)
        {
            _serie = serie;
            AddSchemaNodeOrder(schemaNodeOrder, new string[] { "spPr" });
            Index = GetXmlNodeInt(indexPath);
        }
        internal ExcelChartExDataPoint(ExcelChartExSerie serie, XmlNamespaceManager ns, XmlNode topNode, int index, string[] schemaNodeOrder) : base(ns, topNode)
        {
            _serie = serie;
            AddSchemaNodeOrder(schemaNodeOrder, new string[] { "spPr" });
            Index = index;
        }

        internal const string dataPtPath = "cx:dataPt";
        internal const string SubTotalPath = "cx:layoutPr/cx:subtotals/cx:idx";
        const string indexPath = "@idx";
        /// <summary>
        /// The index of the datapoint
        /// </summary>
        public int Index
        {
            get;
            private set;
        }
        /// <summary>
        /// The data point is a subtotal. Applies for waterfall charts.
        /// </summary>
        public bool SubTotal
        {
            get
            {
                return ExistsNode($"{GetSubTotalPath()}[@val={Index}]");
            }
            set
            {
                string? path = GetSubTotalPath();
                if (value)
                {
                    if (!ExistsNode($"{path}[@val={Index}]"))
                    {
                        XmlElement? idxElement = (XmlElement)CreateNode(path, false, true);
                        idxElement.SetAttribute("val", Index.ToString(CultureInfo.InvariantCulture));
                    }
                }
                else
                {
                    DeleteNode($"{path}/[@val={Index}]");
                }
            }
        }

        private string GetSubTotalPath()
        {
            if(TopNode.LocalName=="series")
            {
                return "cx:layoutPr/cx:subtotals/cx:idx";
            }
            else
            {
                return "../cx:layoutPr/cx:subtotals/cx:idx";
            }
        }

        ExcelDrawingFill _fill = null;
        /// <summary>
        /// A reference to fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    CreateDp();
                    _fill = new ExcelDrawingFill(_serie._chart, NameSpaceManager, TopNode, "cx:spPr", SchemaNodeOrder);
                }
                return _fill;
            }
        }

        ExcelDrawingBorder _line = null;
        /// <summary>
        /// A reference to line properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_line == null)
                {
                    CreateDp();
                    _line = new ExcelDrawingBorder(_serie._chart, NameSpaceManager, TopNode, "cx:spPr/a:ln", SchemaNodeOrder);
                }
                return _line;
            }
        }
        private ExcelDrawingEffectStyle _effect = null;
        /// <summary>
        /// A reference to line properties
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                if (_effect == null)
                {
                    CreateDp();
                    _effect = new ExcelDrawingEffectStyle(_serie._chart, NameSpaceManager, TopNode, "cx:spPr/a:effectLst", SchemaNodeOrder);
                }
                return _effect;
            }
        }
        ExcelDrawing3D _threeD = null;
        /// <summary>
        /// 3D properties
        /// </summary>
        public ExcelDrawing3D ThreeD
        {
            get
            {
                if (_threeD == null)
                {
                    CreateDp();
                    _threeD = new ExcelDrawing3D(NameSpaceManager, TopNode, "cx:spPr", SchemaNodeOrder);
                }
                return _threeD;
            }
        }
        private void CreateDp()
        {
            if (TopNode.LocalName == "series")
            {
                XmlElement pointElement;
                XmlElement? prepend = GetPrependItem();
                if (prepend == null)
                {
                    pointElement = (XmlElement)CreateNode(dataPtPath);
                }
                else
                {
                    pointElement = TopNode.OwnerDocument.CreateElement(dataPtPath, ExcelPackage.schemaChartExMain);
                    prepend.ParentNode.InsertBefore(pointElement, prepend);
                }
                pointElement.SetAttribute("idx", Index.ToString(CultureInfo.InvariantCulture));
                TopNode = pointElement;
            }
        }

        private XmlElement GetPrependItem()
        {
            SortedDictionary<int, ExcelChartExDataPoint>? dic = _serie.DataPoints._dic;
            int prevKey = -1;
            foreach (ExcelChartExDataPoint? v in dic.Values)
            {
                if (v.TopNode.LocalName == "dataPt" && prevKey < v.Index)
                {
                    return (XmlElement)v.TopNode;
                }
            }
            return null;
        }

        void IDrawingStyleBase.CreatespPr()
        {
            base.CreatespPrNode("cx:spPr");
        }
    }
}