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
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A collection of individually formatted datalabels
    /// </summary>
    public class ExcelChartDataLabelCollection : XmlHelper, IEnumerable<ExcelChartDataLabelItem>
    {
        ExcelChart _chart;
        private readonly List<ExcelChartDataLabelItem> _list;
        internal ExcelChartDataLabelCollection(ExcelChart chart, XmlNamespaceManager ns, XmlNode topNode, string[] schemaNodeOrder) : base(ns, topNode)
        {
            this.SchemaNodeOrder = schemaNodeOrder;
            this._list = new List<ExcelChartDataLabelItem>();
            foreach (XmlNode pointNode in this.TopNode.SelectNodes(ExcelChartDataPoint.topNodePath, ns))
            {
                this._list.Add(new ExcelChartDataLabelItem(chart, ns, pointNode, "idx", schemaNodeOrder));
            }

            this._chart = chart;
        }
        /// <summary>
        /// Adds a new chart label to the collection
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns></returns>
        public ExcelChartDataLabelItem Add(int index)
        {
            if (this._list.Count == 0)
            {
                return this.CreateDataLabel(index);
            }
            else
            {
                int ix = this.GetItemAfter(index);
                if (this._list[ix].Index == index)
                {
                    throw (new ArgumentException($"Data label with index {index} already exists"));
                }
                return this.CreateDataLabel(ix);
            }
        }

        private ExcelChartDataLabelItem CreateDataLabel(int idx)
        {
            int pos = this.GetItemAfter(idx);
            XmlElement element = this.CreateElement(idx);
            ExcelChartDataLabelItem? dl = new ExcelChartDataLabelItem(this._chart, this.NameSpaceManager, element, "dLbl", this.SchemaNodeOrder) { Index=idx };

            if (idx < this._list.Count)
            {
                this._list.Insert(idx, dl);
            }
            else
            {
                this._list.Add(dl);
            }

            return dl;
        }

        private XmlElement CreateElement(int idx)
        {
            XmlElement pointNode;
            if (idx < this._list.Count)
            {
                pointNode = this.TopNode.OwnerDocument.CreateElement("c", "dLbl", ExcelPackage.schemaMain);
                this._list[idx].TopNode.InsertBefore(pointNode, this._list[idx].TopNode);
            }
            else
            {
                pointNode = (XmlElement)this.CreateNode("c:dLbl");
            }
            return pointNode;
        }

        private int GetItemAfter(int index)
        {
            for (int i = 0; i < this._list.Count; i++)
            {
                if (index >= this._list[i].Index)
                {
                    return i;
                }
            }
            return this._list.Count;
        }
        /// <summary>
        /// Indexer for the collection
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns></returns>
        public ExcelChartDataLabelItem this[int index]
        {
            get
            {
                return (this._list[index]);
            }
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return this._list.Count;
            }
        }
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ExcelChartDataLabelItem> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this._list.GetEnumerator();
        }
    }
}