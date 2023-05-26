/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/01/2020         EPPlus Software AB       EPPlus 5.3
 *************************************************************************************************/
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Security.Principal;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Slicer
{
    /// <summary>
    /// Tabular data for a pivot table slicer cache.
    /// </summary>
    public class ExcelPivotTableSlicerCacheTabularData : XmlHelper
    {
        const string _topPath = "x14:data/x14:tabular";
        internal readonly ExcelPivotTableSlicerCache _cache;
        internal ExcelPivotTableSlicerCacheTabularData(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTableSlicerCache cache) : base(nsm, topNode)
        {
            this.SchemaNodeOrder = new string[] { "pivotTables", "data" };
            this._cache = cache;
        }
        const string _crossFilterPath = _topPath + "/@crossFilter";
        /// <summary>
        /// How the items that are used in slicer cross filtering are displayed
        /// </summary>
        public eCrossFilter CrossFilter
        {
            get
            {
                return this.GetXmlNodeString(_crossFilterPath).ToEnum(eCrossFilter.ShowItemsWithDataAtTop);
            }
            set
            {
                this.SetXmlNodeString(_crossFilterPath, value.ToEnumString());
            }
        }

        const string _sortOrderPath = _topPath + "/@sortOrder";
        /// <summary>
        /// How the table slicer items are sorted
        /// </summary>
        public eSortOrder SortOrder
        {
            get
            {
                return this.GetXmlNodeString(_sortOrderPath).ToEnum(eSortOrder.Ascending);
            }
            set
            {
                this.SetXmlNodeString(_sortOrderPath, value.ToEnumString());
            }
        }
        const string _customListSortPath = _topPath + "/@customList";

        /// <summary>
        /// If custom lists are used when sorting the items
        /// </summary>
        public bool CustomListSort
        {
            get
            {
                return this.GetXmlNodeBool(_customListSortPath, true);
            }
            set
            {
                this.SetXmlNodeBool(_customListSortPath, value, true);
            }
        }
        const string _showMissingPath = _topPath + "/@showMissing";
        /// <summary>
        /// If the source pivottable has been deleted.
        /// </summary>
        internal bool ShowMissing
        {
            get
            {
                return this.GetXmlNodeBool(_showMissingPath, true);
            }
            set
            {
                this.SetXmlNodeBool(_showMissingPath, value, true);
            }
        }
        private ExcelPivotTableSlicerItemCollection _items =null;
        /// <summary>
        /// The items of the slicer. 
        /// Note that the sort order of this collection is the same as the pivot table field items, not the sortorder of the slicer.
        /// Showing/hiding items are reflects to the pivot table(s) field items collection.
        /// </summary>
        public ExcelPivotTableSlicerItemCollection Items
        {
            get
            {
                if(this._items==null)
                {
                    this._items = new ExcelPivotTableSlicerItemCollection(this._cache);
                }
                return this._items;
            }
        }
        /// <summary>
        /// The pivot table cache id
        /// </summary>
        public int PivotCacheId 
        { 
            get
            {
                return this.GetXmlNodeInt(_topPath + "/@pivotCacheId");
            }
            private set
            {
                this.SetXmlNodeInt(_topPath + "/@pivotCacheId", value);
            } 
        }

        internal void UpdateItemsXml()
        {
            StringBuilder? sb = new StringBuilder();
            int x = 0;
            if (this._cache._field == null)
            {
                return;
            }

            foreach (ExcelPivotTableFieldItem? item in this._cache._field.Items)
            {
                if (item.Type == eItemType.Data)
                {
                    if (item.Hidden)
                    {
                        sb.Append($"<i x=\"{x++}\" />");
                    }
                    else
                    {
                        sb.Append($"<i x=\"{x++}\" s=\"1\"/>");
                    }
                }
            }

            if (this.PivotCacheId < 0)
            {
                this.PivotCacheId = this._cache._field._pivotTable.CacheId;
            }
            XmlElement? dataNode = (XmlElement)this.CreateNode(_topPath+"/x14:items");
            dataNode.SetAttribute("count", x.ToString());
            dataNode.InnerXml = sb.ToString();
        }
    }
}