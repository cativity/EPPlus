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

namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// A collection of filter columns for an autofilter of table in a worksheet
    /// </summary>
    public class ExcelFilterColumnCollection : XmlHelper, IEnumerable<ExcelFilterColumn>
    {
        SortedDictionary<int, ExcelFilterColumn> _columns = new SortedDictionary<int, ExcelFilterColumn>();
        ExcelAutoFilter _autoFilter;
        internal ExcelFilterColumnCollection(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelAutoFilter autofilter) : base(namespaceManager, topNode)
        {
            this._autoFilter = autofilter;
            foreach (XmlElement node in topNode.SelectNodes("d:filterColumn", namespaceManager))
            {
                if(!int.TryParse(node.Attributes["colId"].Value, out int position))
                {
                    throw (new Exception("Invalid filter. Missing colId on filterColumn"));
                }
                switch (node.FirstChild?.Name)
                {
                    case "filters":
                        this._columns.Add(position, new ExcelValueFilterColumn(namespaceManager, node));
                        break;
                    case "customFilters":
                        this._columns.Add(position, new ExcelCustomFilterColumn(namespaceManager, node));
                        break;
                    case "colorFilter":
                        this._columns.Add(position, new ExcelColorFilterColumn(namespaceManager, node));
                        break;
                    case "iconFilter":
                        this._columns.Add(position, new ExcelIconFilterColumn(namespaceManager, node));
                        break;
                    case "dynamicFilter":
                        this._columns.Add(position, new ExcelDynamicFilterColumn(namespaceManager, node));
                        break;
                    case "top10":
                        this._columns.Add(position, new ExcelTop10FilterColumn(namespaceManager, node));
                        break;
                }
            }
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return this._columns.Count;
            }
        }
        internal XmlNode Add(int position, string topNodeName)
        {
            XmlElement node;
            if (position >= this._autoFilter.Address.Columns)
            {
                throw (new ArgumentOutOfRangeException("Position is outside of the range"));
            }
            if (this._columns.ContainsKey(position))
            {
                throw (new ArgumentOutOfRangeException("Position already exists"));
            }
            foreach (ExcelFilterColumn? c in this._columns.Values)
            {
                if (c.Position > position)
                {
                    node = this.GetColumnNode(position, topNodeName);
                    return c.TopNode.ParentNode.InsertBefore(node, c.TopNode);
                }
            }
            node = this.GetColumnNode(position, topNodeName);
            return this.TopNode.AppendChild(node);
        }

        private XmlElement GetColumnNode(int position, string topNodeName)
        {
            XmlElement node = this.TopNode.OwnerDocument.CreateElement("filterColumn", ExcelPackage.schemaMain);
            node.SetAttribute("colId", position.ToString());
            XmlElement? subNode = this.TopNode.OwnerDocument.CreateElement(topNodeName, ExcelPackage.schemaMain);
            node.AppendChild(subNode);
            return node;
        }
        /// <summary>
        /// Indexer of filtercolumns
        /// </summary>
        /// <param name="index">The column index starting from zero</param>
        /// <returns>A filter column</returns>
        public ExcelFilterColumn this[int index]
        {
            get
            {
                if(this._columns.ContainsKey(index))
                {
                    return this._columns[index];
                }
                else
                {
                    return null;
                }
            }
        }
        /// <summary>
        /// Adds a value filter for the specified column position
        /// </summary>
        /// <param name="position">The column position</param>
        /// <returns>The value filter</returns>
        public ExcelValueFilterColumn AddValueFilterColumn(int position)
        {
            XmlNode? node = this.Add(position, "filters");
            ExcelValueFilterColumn? col = new ExcelValueFilterColumn(this.NameSpaceManager, node);
            this._columns.Add(position, col);
            return col;
        }
        /// <summary>
        /// Adds a custom filter for the specified column position
        /// </summary>
        /// <param name="position">The column position</param>
        /// <returns>The custom filter</returns>
        public ExcelCustomFilterColumn AddCustomFilterColumn(int position)
        {
            XmlNode? node = this.Add(position, "customFilters");
            ExcelCustomFilterColumn? col= new ExcelCustomFilterColumn(this.NameSpaceManager, node);
            this._columns.Add(position, col);
            return col;
        }
        /// <summary>
        /// Adds a color filter for the specified column position
        /// Note: EPPlus doesn't filter color filters when <c>ApplyFilter</c> is called.
        /// </summary>
        /// <param name="position">The column position</param>
        /// <returns>The color filter</returns>
        public ExcelColorFilterColumn AddColorFilterColumn(int position)
        {
            XmlNode? node = this.Add(position, "colorFilter");
            ExcelColorFilterColumn? col = new ExcelColorFilterColumn(this.NameSpaceManager, node);
            this._columns.Add(position, col);
            return col;
        }
        /// <summary>
        /// Adds a icon filter for the specified column position
        /// Note: EPPlus doesn't filter icon filters when <c>ApplyFilter</c> is called.
        /// </summary>
        /// <param name="position">The column position</param>
        /// <returns>The color filter</returns>
        public ExcelIconFilterColumn AddIconFilterColumn(int position)
        {
            XmlNode? node = this.Add(position, "iconFilter");
            ExcelIconFilterColumn? col = new ExcelIconFilterColumn(this.NameSpaceManager, node);
            this._columns.Add(position, col);
            return col;
        }
        /// <summary>
        /// Adds a top10 filter for the specified column position
        /// </summary>
        /// <param name="position">The column position</param>
        /// <returns>The top 10 filter</returns>
        public ExcelTop10FilterColumn AddTop10FilterColumn(int position)
        {
            XmlNode? node = this.Add(position, "top10");
            ExcelTop10FilterColumn? col = new ExcelTop10FilterColumn(this.NameSpaceManager, node);
            this._columns.Add(position, col);
            return col;
        }
        /// <summary>
        /// Adds a dynamic filter for the specified column position
        /// </summary>
        /// <param name="position">The column position</param>
        /// <returns>The dynamic filter</returns>
        public ExcelDynamicFilterColumn AddDynamicFilterColumn(int position)
        {
            XmlNode? node = this.Add(position, "dynamicFilter");
            ExcelDynamicFilterColumn? col = new ExcelDynamicFilterColumn(this.NameSpaceManager, node);
            this._columns.Add(position, col);
            return col;
        }
        /// <summary>
        /// Gets the enumerator of the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ExcelFilterColumn> GetEnumerator()
        {
            return this._columns.Values.GetEnumerator();
        }
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this._columns.Values.GetEnumerator();
        }
        /// <summary>
        /// Remove the filter column with the position from the collection
        /// </summary>
        /// <param name="position">The index of the column to remove</param>
        public void RemoveAt(int position)
        {
            if(!this._columns.ContainsKey(position))
            {
                throw new InvalidOperationException($"Column with position {position} does not exist in the filter collection");
            }

            this.Remove(this._columns[position]);
        }
        /// <summary>
        /// Remove the filter column from the collection
        /// </summary>
        /// <param name="column">The column</param>
        public void Remove(ExcelFilterColumn column)
        {
            XmlNode? node = column.TopNode;
            node.ParentNode.RemoveChild(node);
            this._columns.Remove(column.Position);
        }
    }
}