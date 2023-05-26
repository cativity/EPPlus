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
using System.Xml;
using System.Globalization;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System.Linq;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Drawing;
using System.Text;
using System.Collections;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.Core;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Table.PivotTable.Filter;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A pivot table field.
    /// </summary>
    public class ExcelPivotTableField : XmlHelper
    {
        internal ExcelPivotTable _pivotTable;
        internal ExcelPivotTableCacheField _cacheField = null;        
        internal ExcelPivotTableField(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTable table, int index, int baseIndex) :
            base(ns, topNode)
        {
            this.SchemaNodeOrder = new string[] { "items","autoSortScope" };
            this.Index = index;
            this.BaseIndex = baseIndex;
            this._pivotTable = table;
            if(this.NumFmtId.HasValue)
            {
                ExcelStyles? styles = table.WorkSheet.Workbook.Styles;
                int ix = styles.NumberFormats.FindIndexById(this.NumFmtId.Value.ToString(CultureInfo.InvariantCulture));
                if(ix>=0)
                {
                    this.Format = styles.NumberFormats[ix].Format;
                }
            }
        }
        /// <summary>
        /// The index of the pivot table field
        /// </summary>
        public int Index
        {
            get;
            set;
        }
        /// <summary>
        /// The base line index of the pivot table field
        /// </summary>
        internal int BaseIndex
        {
            get;
            set;
        }
        /// <summary>
        /// Name of the field
        /// </summary>
        public string Name
        {
            get
            {
                string v = this.GetXmlNodeString("@name");
                if (v == "")
                {
                    return this._cacheField?.Name;
                }
                else
                {
                    return v;
                }
            }
            set
            {
                this.SetXmlNodeString("@name", value);
            }
        }
        /// <summary>
        /// Compact mode
        /// </summary>
        public bool Compact
        {
            get
            {
                return this.GetXmlNodeBool("@compact");
            }
            set
            {
                this.SetXmlNodeBool("@compact", value);
            }
        }
        /// <summary>
        /// A boolean that indicates whether the items in this field should be shown in Outline form
        /// </summary>
        public bool Outline
        {
            get
            {
                return this.GetXmlNodeBool("@outline");
            }
            set
            {
                this.SetXmlNodeBool("@outline", value);
            }
        }
        /// <summary>
        /// The custom text that is displayed for the subtotals label
        /// </summary>
        public bool SubtotalTop
        {
            get
            {
                return this.GetXmlNodeBool("@subtotalTop");
            }
            set
            {
                this.SetXmlNodeBool("@subtotalTop", value);
            }
        }
        /// <summary>
        /// Indicates whether the field can have multiple items selected in the page field
        /// </summary>
        public bool MultipleItemSelectionAllowed
        {
            get
            {
                return this.GetXmlNodeBool("@multipleItemSelectionAllowed");
            }
            set
            {
                this.SetXmlNodeBool("@multipleItemSelectionAllowed", value);
                if(value && this.IsPageField)
                {
                    this.PageFieldSettings.SelectedItem = -1;
                }                
            }
        }
        #region Show properties
        /// <summary>
        /// Indicates whether to show all items for this field
        /// </summary>
        public bool ShowAll
        {
            get
            {
                return this.GetXmlNodeBool("@showAll");
            }
            set
            {
                this.SetXmlNodeBool("@showAll", value);
            }
        }
        /// <summary>
        /// Indicates whether to hide drop down buttons on PivotField headers
        /// </summary>
        public bool ShowDropDowns
        {
            get
            {
                return this.GetXmlNodeBool("@showDropDowns");
            }
            set
            {
                this.SetXmlNodeBool("@showDropDowns", value);
            }
        }
        /// <summary>
        /// Indicates whether this hierarchy is omitted from the field list
        /// </summary>
        public bool ShowInFieldList
        {
            get
            {
                return this.GetXmlNodeBool("@showInFieldList");
            }
            set
            {
                this.SetXmlNodeBool("@showInFieldList", value);
            }
        }
        /// <summary>
        /// Indicates whether to show the property as a member caption
        /// </summary>
        public bool ShowAsCaption
        {
            get
            {
                return this.GetXmlNodeBool("@showPropAsCaption");
            }
            set
            {
                this.SetXmlNodeBool("@showPropAsCaption", value);
            }
        }
        /// <summary>
        /// Indicates whether to show the member property value in a PivotTable cell
        /// </summary>
        public bool ShowMemberPropertyInCell
        {
            get
            {
                return this.GetXmlNodeBool("@showPropCell");
            }
            set
            {
                this.SetXmlNodeBool("@showPropCell", value);
            }
        }
        /// <summary>
        /// Indicates whether to show the member property value in a tooltip on the appropriate PivotTable cells
        /// </summary>
        public bool ShowMemberPropertyToolTip
        {
            get
            {
                return this.GetXmlNodeBool("@showPropTip");
            }
            set
            {
                this.SetXmlNodeBool("@showPropTip", value);
            }
        }
        #endregion
        /// <summary>
        /// The type of sort that is applied to this field
        /// </summary>
        public eSortType Sort
        {
            get
            {
                string v = this.GetXmlNodeString("@sortType");
                return v == "" ? eSortType.None : (eSortType)Enum.Parse(typeof(eSortType), v, true);
            }
            set
            {
                if (value == eSortType.None)
                {
                    this.DeleteNode("@sortType");
                }
                else
                {
                    this.SetXmlNodeString("@sortType", value.ToString().ToLower(CultureInfo.InvariantCulture));
                }
            }
        }

        /// <summary>
        /// Set auto sort on a data field for this field.
        /// </summary>
        /// <param name="dataField">The data field to sort on</param>
        /// <param name="sortType">Sort ascending or descending</param>
        public void SetAutoSort(ExcelPivotTableDataField dataField, eSortType sortType=eSortType.Ascending)
        {
            if(dataField.Field._pivotTable!= this._pivotTable)
            {
                throw (new ArgumentException("The dataField is from another pivot table"));
            }

            this.Sort = sortType;
            XmlNode? node = this.CreateNode("d:autoSortScope/d:pivotArea");
            if (this.AutoSort == null)
            {
                this.AutoSort = new ExcelPivotAreaAutoSort(this.NameSpaceManager, node, this._pivotTable);
                this.AutoSort.FieldPosition = 0;
                this.AutoSort.Outline = false;
                this.AutoSort.DataOnly = false;
            }

            this.AutoSort.DeleteNode("d:references");
            this.AutoSort.Conditions.Fields.Clear();
            this.AutoSort.Conditions.DataFields.Clear();
            this.AutoSort.Conditions.DataFields.Add(dataField);
        }
        /// <summary>
        /// Remove auto sort and set the <see cref="AutoSort"/> property to null
        /// </summary>
        public void RemoveAutoSort()
        {
            if (this.AutoSort !=null)
            {
                this.AutoSort.DeleteNode("d:autoSortScope");
                this.AutoSort = null;
            }
        }

        /// <summary>
        /// Auto sort for a field. Sort is set on a data field for a row/column field.
        /// Use <see cref="SetAutoSort(ExcelPivotTableDataField, eSortType)"/> to set auto sort 
        /// Use <seealso cref="RemoveAutoSort"/> to remove auto sort and set this property to null
        /// </summary>
        public ExcelPivotAreaAutoSort AutoSort
        {
            get;
            private set;
        }
        /// <summary>
        /// A boolean that indicates whether manual filter is in inclusive mode
        /// </summary>
        public bool IncludeNewItemsInFilter
        {
            get
            {
                return this.GetXmlNodeBool("@includeNewItemsInFilter");
            }
            set
            {
                this.SetXmlNodeBool("@includeNewItemsInFilter", value);
            }
        }
        /// <summary>
        /// Enumeration of the different subtotal operations that can be applied to page, row or column fields
        /// </summary>
        public eSubTotalFunctions SubTotalFunctions
        {
            get
            {
                eSubTotalFunctions ret = 0;
                XmlNodeList nl = this.TopNode.SelectNodes("d:items/d:item/@t", this.NameSpaceManager);
                if (nl.Count == 0)
                {
                    return eSubTotalFunctions.None;
                }

                foreach (XmlAttribute item in nl)
                {
                    try
                    {
                        ret |= (eSubTotalFunctions)Enum.Parse(typeof(eSubTotalFunctions), item.Value, true);
                    }
                    catch (ArgumentException ex)
                    {
                        throw new ArgumentException("Unable to parse value of " + item.Value + " to a valid pivot table subtotal function", ex);
                    }
                }
                return ret;
            }
            set
            {
                if ((value & eSubTotalFunctions.None) == eSubTotalFunctions.None && (value != eSubTotalFunctions.None))
                {
                    throw (new ArgumentException("Value None cannot be combined with other values."));
                }
                if ((value & eSubTotalFunctions.Default) == eSubTotalFunctions.Default && (value != eSubTotalFunctions.Default))
                {
                    throw (new ArgumentException("Value Default cannot be combined with other values."));
                }


                // remove old attribute                 
                XmlNodeList nl = this.TopNode.SelectNodes("d:items/d:item/@t", this.NameSpaceManager);
                if (nl.Count > 0)
                {
                    foreach (XmlAttribute item in nl)
                    {
                        this.DeleteNode("@" + item.Value + "Subtotal");
                        item.OwnerElement.ParentNode.RemoveChild(item.OwnerElement);
                    }
                }


                if (value == eSubTotalFunctions.None)
                {
                    // for no subtotals, set defaultSubtotal to off
                    this.SetXmlNodeBool("@defaultSubtotal", false);
                    //TopNode.InnerXml = "<items count=\"1\"><item x=\"0\"/></items>";
                    //_cacheFieldHelper.TopNode.InnerXml = "<sharedItems count=\"1\"><m/></sharedItems>";
                }
                else
                {
                    string innerXml = "";
                    int count = 0;
                    foreach (eSubTotalFunctions e in Enum.GetValues(typeof(eSubTotalFunctions)))
                    {
                        if ((value & e) == e)
                        {
                            string? newTotalType = e.ToString();
                            string? totalType = char.ToLowerInvariant(newTotalType[0]) + newTotalType.Substring(1);
                            // add new attribute
                            this.SetXmlNodeBool("@" + totalType + "Subtotal", true);
                            innerXml += "<item t=\"" + totalType + "\" />";
                            count++;
                        }
                    }

                    this.SetXmlNodeInt("d:items/@count", count);
                    XmlNode? itemsNode= this.GetNode("d:items");
                    itemsNode.InnerXml = innerXml;
                }
            }
        }
        /// <summary>
        /// Type of axis
        /// </summary>
        public ePivotFieldAxis Axis
        {
            get
            {
                switch (this.GetXmlNodeString("@axis"))
                {
                    case "axisRow":
                        return ePivotFieldAxis.Row;
                    case "axisCol":
                        return ePivotFieldAxis.Column;
                    case "axisPage":
                        return ePivotFieldAxis.Page;
                    case "axisValues":
                        return ePivotFieldAxis.Values;
                    default:
                        return ePivotFieldAxis.None;
                }
            }
            internal set
            {
                switch (value)
                {
                    case ePivotFieldAxis.Row:
                        this.SetXmlNodeString("@axis", "axisRow");
                        break;
                    case ePivotFieldAxis.Column:
                        this.SetXmlNodeString("@axis", "axisCol");
                        break;
                    case ePivotFieldAxis.Values:
                        this.SetXmlNodeString("@axis", "axisValues");
                        break;
                    case ePivotFieldAxis.Page:
                        this.SetXmlNodeString("@axis", "axisPage");
                        break;
                    default:
                        this.DeleteNode("@axis");
                        break;
                }
            }
        }
        /// <summary>
        /// If the field is a row field
        /// </summary>
        public bool IsRowField
        {
            get
            {
                return (this.TopNode.SelectSingleNode(string.Format("../../d:rowFields/d:field[@x={0}]", this.Index), this.NameSpaceManager) != null);
            }
            internal set
            {
                if (value)
                {
                    XmlNode? rowsNode = this.TopNode.SelectSingleNode("../../d:rowFields", this.NameSpaceManager);
                    if (rowsNode == null)
                    {
                        this._pivotTable.CreateNode("d:rowFields");
                    }
                    rowsNode = this.TopNode.SelectSingleNode("../../d:rowFields", this.NameSpaceManager);

                    AppendField(rowsNode, this.Index, "field", "x");
                    if (this.Grouping == null)
                    {
                        if (this.BaseIndex == this.Index)
                        {
                            this.TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";
                        }
                        else
                        {
                            this.TopNode.InnerXml = "<items count=\"0\"/>";
                        }
                    }
                }
                else
                {
                    XmlElement node = this.TopNode.SelectSingleNode(string.Format("../../d:rowFields/d:field[@x={0}]", this.Index), this.NameSpaceManager) as XmlElement;
                    if (node != null)
                    {
                        node.ParentNode.RemoveChild(node);
                    }
                }
            }
        }
        /// <summary>
        /// If the field is a column field
        /// </summary>
        public bool IsColumnField
        {
            get
            {
                return (this.TopNode.SelectSingleNode(string.Format("../../d:colFields/d:field[@x={0}]", this.Index), this.NameSpaceManager) != null);
            }
            internal set
            {
                if (value)
                {
                    XmlNode? columnsNode = this.TopNode.SelectSingleNode("../../d:colFields", this.NameSpaceManager);
                    if (columnsNode == null)
                    {
                        this._pivotTable.CreateNode("d:colFields");
                    }
                    columnsNode = this.TopNode.SelectSingleNode("../../d:colFields", this.NameSpaceManager);

                    AppendField(columnsNode, this.Index, "field", "x");
                    if (this.BaseIndex == this.Index)
                    {
                        this.TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";
                    }
                    else
                    {
                        this.TopNode.InnerXml = "<items count=\"0\"></items>";
                    }
                }
                else
                {
                    XmlElement node = this.TopNode.SelectSingleNode(string.Format("../../d:colFields/d:field[@x={0}]", this.Index), this.NameSpaceManager) as XmlElement;
                    if (node != null)
                    {
                        node.ParentNode.RemoveChild(node);
                    }
                }
            }
        }
        /// <summary>
        /// If the field is a datafield
        /// </summary>
        public bool IsDataField
        {
            get
            {
                return this.GetXmlNodeBool("@dataField", false);
            }
            set
            {
                this.SetXmlNodeBool("@dataField", value, false);
            }
        }
        /// <summary>
        /// If the field is a page field.
        /// </summary>
        public bool IsPageField
        {
            get
            {
                return (this.Axis == ePivotFieldAxis.Page);
            }
            internal set
            {
                if (value)
                {
                    XmlNode? dataFieldsNode = this.TopNode.SelectSingleNode("../../d:pageFields", this.NameSpaceManager);
                    if (dataFieldsNode == null)
                    {
                        this._pivotTable.CreateNode("d:pageFields");
                        dataFieldsNode = this.TopNode.SelectSingleNode("../../d:pageFields", this.NameSpaceManager);
                    }

                    this.TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";

                    XmlElement node = AppendField(dataFieldsNode, this.Index, "pageField", "fld");
                    this._pageFieldSettings = new ExcelPivotTablePageFieldSettings(this.NameSpaceManager, node, this, this.Index);
                }
                else
                {
                    this._pageFieldSettings = null;
                    XmlElement node = this.TopNode.SelectSingleNode(string.Format("../../d:pageFields/d:pageField[@fld={0}]", this.Index), this.NameSpaceManager) as XmlElement;
                    if (node != null)
                    {
                        node.ParentNode.RemoveChild(node);
                    }
                }
            }
        }
        //public ExcelPivotGrouping DateGrouping
        //{

        //}
        internal ExcelPivotTablePageFieldSettings _pageFieldSettings = null;
        /// <summary>
        /// Page field settings
        /// </summary>
        public ExcelPivotTablePageFieldSettings PageFieldSettings
        {
            get
            {
                return this._pageFieldSettings;
            }
        }
        /// <summary>
        /// Date group by
        /// </summary>
        internal eDateGroupBy DateGrouping
        {
            get
            {
                return this.Cache.DateGrouping;
            }
        }
        /// <summary>
        /// Grouping settings. 
        /// Null if the field has no grouping otherwise ExcelPivotTableFieldDateGroup or ExcelPivotTableFieldNumericGroup.
        /// </summary>        
        public ExcelPivotTableFieldGroup Grouping
        {
            get
            {                
                return this.Cache.Grouping;
            }
        }
        /// <summary>
        /// The numberformat to use for the column
        /// </summary>
        public string Format { get; set; }
        #region Private & internal Methods
        internal static XmlElement AppendField(XmlNode rowsNode, int index, string fieldNodeText, string indexAttrText)
        {
            XmlElement prevField = null, newElement;
            foreach (XmlElement field in rowsNode.ChildNodes)
            {
                string x = field.GetAttribute(indexAttrText);
                int fieldIndex;
                if (int.TryParse(x, out fieldIndex))
                {
                    if (fieldIndex == index)    //Row already exists
                    {
                        return field;
                    }
                }
                prevField = field;
            }
            newElement = rowsNode.OwnerDocument.CreateElement(fieldNodeText, ExcelPackage.schemaMain);
            newElement.SetAttribute(indexAttrText, index.ToString());
            rowsNode.InsertAfter(newElement, prevField);

            return newElement;
        }
        #endregion
        internal ExcelPivotTableFieldItemsCollection _items = null;
        /// <summary>
        /// Pivottable field Items. Used for grouping.
        /// </summary>
        public ExcelPivotTableFieldItemsCollection Items
        {
            get
            {
                if (this._items == null)
                {
                    this.LoadItems();
                }
                return this._items;
            }
        }

        internal void LoadItems()
        {
            this._items = new ExcelPivotTableFieldItemsCollection(this);
            if (this.Cache.DatabaseField == false && (this.IsColumnField == false && this.IsRowField == false && this.IsRowField == false))
            {
                return;
            }

            EPPlusReadOnlyList<object> cacheItems;
            if (this.Cache.Grouping == null)
            {
                cacheItems = this.Cache.SharedItems;
            }
            else
            {
                cacheItems = this.Cache.GroupItems;
            }

            foreach (XmlElement node in this.TopNode.SelectNodes("d:items//d:item", this.NameSpaceManager))
            {
                ExcelPivotTableFieldItem? item = new ExcelPivotTableFieldItem(node);
                if (item.X >= 0 && item.X < cacheItems.Count)
                {
                    item.Value = cacheItems[item.X];
                }

                this._items.AddInternal(item);
            }
        }
        /// <summary>
        /// A reference to the cache for the pivot table field.
        /// </summary>
        public ExcelPivotTableCacheField Cache
        {
            get
            {
                return this._pivotTable.CacheDefinition._cacheReference.Fields[this.Index];
            }
        }
        /// <summary>
        /// Add numberic grouping to the field
        /// </summary>
        /// <param name="Start">Start value</param>
        /// <param name="End">End value</param>
        /// <param name="Interval">Interval</param>
        public void AddNumericGrouping(double Start, double End, double Interval)
        {
            this.ValidateGrouping();
            this._cacheField.SetNumericGroup(this.BaseIndex, Start, End, Interval);
            this.UpdateGroupItems(this._cacheField, true);
            UpdatePivotTableGroupItems(this, this._pivotTable.CacheDefinition._cacheReference, true);
        }
        /// <summary>
        /// Will add a slicer to the pivot table field
        /// </summary>
        /// <returns>The <see cref="ExcelPivotTableSlicer">Slicer</see>/></returns>
        public ExcelPivotTableSlicer AddSlicer()
        {
            if (this._slicer != null)
            {
                throw new InvalidOperationException("");
            }

            this._slicer = this._pivotTable.WorkSheet.Drawings.AddPivotTableSlicer(this);
            return this._slicer;
        }
        ExcelPivotTableSlicer _slicer = null;
        /// <summary>
        /// A slicer attached to the pivot table field.
        /// If the field has multiple slicers attached, the first slicer will be returned.
        /// </summary>
        public ExcelPivotTableSlicer Slicer
        {
            get 
            {
                if (this._slicer == null && this._pivotTable.WorkSheet.Workbook.ExistsNode($"d:extLst/d:ext[@uri='{ExtLstUris.WorkbookSlicerPivotTableUri}']"))
                {
                    foreach (ExcelWorksheet? ws in this._pivotTable.WorkSheet.Workbook.Worksheets)
                    {
                        foreach (ExcelDrawing? d in ws.Drawings)
                        {
                            if (d is ExcelPivotTableSlicer s && s.Cache != null && s.Cache.PivotTables.Contains(this._pivotTable) && this.Index==s.Cache._field.Index)
                            {
                                this._slicer = s;
                                return this._slicer;
                            }
                        }
                    }
                }
                return this._slicer;
            }
            internal set
            {
                this._slicer = value;
            }
        }
        /// <summary>
        /// Add a date grouping on this field.
        /// </summary>
        /// <param name="groupBy">Group by</param>
        public void AddDateGrouping(eDateGroupBy groupBy)
        {
            this.AddDateGrouping(groupBy, DateTime.MinValue, DateTime.MaxValue, 1);
        }
        /// <summary>
        /// Add a date grouping on this field.
        /// </summary>
        /// <param name="groupBy">Group by</param>
        /// <param name="startDate">Fixed start date. Use DateTime.MinValue for auto</param>
        /// <param name="endDate">Fixed end date. Use DateTime.MaxValue for auto</param>
        public void AddDateGrouping(eDateGroupBy groupBy, DateTime startDate, DateTime endDate)
        {
            this.AddDateGrouping(groupBy, startDate, endDate, 1);
        }
        /// <summary>
        /// Add a date grouping on this field.
        /// </summary>
        /// <param name="days">Number of days when grouping on days</param>
        /// <param name="startDate">Fixed start date. Use DateTime.MinValue for auto</param>
        /// <param name="endDate">Fixed end date. Use DateTime.MaxValue for auto</param>
        public void AddDateGrouping(int days, DateTime startDate, DateTime endDate)
        {
            this.AddDateGrouping(eDateGroupBy.Days, startDate, endDate, days);
        }
        private ExcelPivotTableField AddField(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, ref bool firstField)
        {
            return this.AddField(groupBy, startDate, endDate, ref firstField, 1);
        }
        private ExcelPivotTableField AddField(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, ref bool firstField, int interval)
        {
            if (firstField == false)
            {
                ExcelPivotTableField field = this._pivotTable.Fields.AddDateGroupField(this.Index);

                XmlNode rowColFields;
                if (this.IsRowField)
                {
                    rowColFields = this.TopNode.SelectSingleNode("../../d:rowFields", this.NameSpaceManager);
                }
                else
                {
                    rowColFields = this.TopNode.SelectSingleNode("../../d:colFields", this.NameSpaceManager);
                }

                int index = 0;
                foreach (XmlElement rowfield in rowColFields.ChildNodes)
                {
                    if (int.TryParse(rowfield.GetAttribute("x"), out int fieldIndex))
                    {
                        if (this._pivotTable.Fields[fieldIndex].BaseIndex == this.BaseIndex)
                        {
                            XmlElement? newElement = rowColFields.OwnerDocument.CreateElement("field", ExcelPackage.schemaMain);
                            newElement.SetAttribute("x", field.Index.ToString());
                            rowColFields.InsertBefore(newElement, rowfield);
                            break;
                        }
                    }
                    index++;
                }

                PivotTableCacheInternal? cacheRef = this._pivotTable.CacheDefinition._cacheReference;
                field._cacheField = cacheRef.AddDateGroupField(field, groupBy, startDate, endDate, interval);
                UpdatePivotTableGroupItems(field, cacheRef, false);

                if (this.IsRowField)
                {
                    this._pivotTable.RowFields.Insert(field, index);
                }
                else
                {
                    this._pivotTable.ColumnFields.Insert(field, index);
                }

                return field;
            }
            else
            {
                firstField = false;
                this.Compact = false;
                this._cacheField.SetDateGroup(this, groupBy, startDate, endDate, interval);
                UpdatePivotTableGroupItems(this, this._pivotTable.CacheDefinition._cacheReference, true);
                return this;
            }
        }
        private static void UpdatePivotTableGroupItems(ExcelPivotTableField field, PivotTableCacheInternal cacheRef, bool addTypeDefault)
        {
            foreach (ExcelPivotTable? pt in cacheRef._pivotTables)
            {
                ExcelPivotTableCacheField? f = cacheRef.Fields[field.Index];
                if (f.Grouping is ExcelPivotTableFieldDateGroup)
                {
                    if(field.Index >= pt.Fields.Count)
                    {
                         ExcelPivotTableField? newField = pt.Fields.AddDateGroupField((int)f.Grouping.BaseIndex);
                        newField._cacheField = f;
                    }

                    pt.Fields[field.Index].UpdateGroupItems(f, addTypeDefault);
                }
                else
                { 
                    pt.Fields[field.Index].UpdateGroupItems(f, addTypeDefault);
                }
            }
        }

        internal void UpdateGroupItems(ExcelPivotTableCacheField cacheField, bool addTypeDefault)
        {
            XmlElement itemsNode = this.CreateNode("d:items") as XmlElement;
            this._items = new ExcelPivotTableFieldItemsCollection(this);
            itemsNode.RemoveAll();
            for (int x = 0; x < cacheField.GroupItems.Count; x++)
            {
                this._items.AddInternal(new ExcelPivotTableFieldItem() { X = x, Value=cacheField.GroupItems[x] });
            }
            if(addTypeDefault)
            {
                this._items.AddInternal(new ExcelPivotTableFieldItem() { Type = eItemType.Default});
            }
        }
        private void AddDateGrouping(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, int groupInterval)
        {
            if (groupInterval < 1 || groupInterval >= Int16.MaxValue)
            {
                throw (new ArgumentOutOfRangeException("Group interval is out of range"));
            }
            if (groupInterval > 1 && groupBy != eDateGroupBy.Days)
            {
                throw (new ArgumentException("Group interval is can only be used when groupBy is Days"));
            }
            if(this.Cache.DatabaseField==false)
            {
                throw new InvalidOperationException("The field for grouping cannot be a calculated field.");
            }

            this.ValidateGrouping();

            this._items = null;

            bool firstField = true;
            int fields = this._pivotTable.Fields.Count;
            //Seconds
            if ((groupBy & eDateGroupBy.Seconds) == eDateGroupBy.Seconds)
            {
                this.AddField(eDateGroupBy.Seconds, startDate, endDate, ref firstField);
            }
            //Minutes
            if ((groupBy & eDateGroupBy.Minutes) == eDateGroupBy.Minutes)
            {
                this.AddField(eDateGroupBy.Minutes, startDate, endDate, ref firstField);
            }
            //Hours
            if ((groupBy & eDateGroupBy.Hours) == eDateGroupBy.Hours)
            {
                this.AddField(eDateGroupBy.Hours, startDate, endDate, ref firstField);
            }
            //Days
            if ((groupBy & eDateGroupBy.Days) == eDateGroupBy.Days)
            {
                this.AddField(eDateGroupBy.Days, startDate, endDate, ref firstField, groupInterval);
            }
            //Month
            if ((groupBy & eDateGroupBy.Months) == eDateGroupBy.Months)
            {
                this.AddField(eDateGroupBy.Months, startDate, endDate, ref firstField);
            }
            //Quarters
            if ((groupBy & eDateGroupBy.Quarters) == eDateGroupBy.Quarters)
            {
                this.AddField(eDateGroupBy.Quarters, startDate, endDate, ref firstField);
            }
            //Years
            if ((groupBy & eDateGroupBy.Years) == eDateGroupBy.Years)
            {
                this.AddField(eDateGroupBy.Years, startDate, endDate, ref firstField);
            }

            if (fields> this._pivotTable.Fields.Count)
            {
                this._cacheField.SetXmlNodeString("d:fieldGroup/@par", (this._pivotTable.Fields.Count-1).ToString());
            }

            if (groupInterval != 1)
            {
                this._cacheField.SetXmlNodeString("d:fieldGroup/d:rangePr/@groupInterval", groupInterval.ToString());
            }
            else
            {
                this._cacheField.DeleteNode("d:fieldGroup/d:rangePr/@groupInterval");
            }
        }

        private void ValidateGrouping()
        {
            if (this.Cache.DatabaseField == false)
            {
                throw new InvalidOperationException("The field for grouping cannot be a calculated field.");
            }

            if (!(this.IsColumnField || this.IsRowField))
            {
                throw (new Exception("Field must be a row or column field"));
            }
            foreach (ExcelPivotTableField? field in this._pivotTable.Fields)
            {
                if (field.Grouping != null)
                {
                    throw (new Exception("Grouping already exists"));
                }
            }
        }
        internal string SaveToXml()
        {
            StringBuilder? sb = new StringBuilder();
            Dictionary<object, int>? cacheLookup = this._pivotTable.CacheDefinition._cacheReference.Fields[this.Index]._cacheLookup;
            if(this.AutoSort!=null)
            {
                this.AutoSort.Conditions.UpdateXml();
            }
            if (cacheLookup == null)
            {
                return "";
            }

            if (cacheLookup.Count==0)
            {
                this.DeleteNode("d:items");       //Creates or return the existing node
            }
            else if (this.Items.Count > 0)
            {
                int hasMultipleSelectedCount = 0;
                foreach (ExcelPivotTableFieldItem? item in this.Items)
                {
                    object? v = item.Value ?? ExcelPivotTable.PivotNullValue;
                    if (item.Type==eItemType.Data && cacheLookup.TryGetValue(v, out int x))
                    {
                        item.X = cacheLookup[v];
                    }
                    else
                    {
                        item.X = -1;
                    }
                    if (hasMultipleSelectedCount<=1 && item.Hidden==false && item.Type!=eItemType.Default)
                    {
                        hasMultipleSelectedCount++;
                    }

                    item.GetXmlString(sb);
                }
                if (hasMultipleSelectedCount > 1 && this.IsPageField)
                {
                    this.PageFieldSettings.SelectedItem = -1;
                }

                XmlElement? node = (XmlElement)this.CreateNode("d:items");       //Creates or return the existing node
                node.InnerXml = sb.ToString();
                node.SetAttribute("count", this.Items.Count.ToString());
            }

            return sb.ToString();
        }
        ExcelPivotTableFieldFilterCollection _filters = null;
        /// <summary>
        /// Filters used on the pivot table field.
        /// </summary>
        public ExcelPivotTableFieldFilterCollection Filters
        {
            get { return this._filters ??= new ExcelPivotTableFieldFilterCollection(this); }
        }

        internal int? NumFmtId 
        {
            get
            {
                return this.GetXmlNodeIntNull("@numFmtId");
            }
            set
            {
                this.SetXmlNodeInt("@numFmtId", value);
            }
        }

        /// <summary>
        /// Allow as column field?
        /// </summary>
        internal bool DragToCol 
        { 
            get
            {
                return this.GetXmlNodeBool("@dragToCol", true);
            }
        }
        /// <summary>
        /// Allow as page row?
        /// </summary>
        internal bool DragToRow
        {
            get
            {
                return this.GetXmlNodeBool("@dragToRow", true);
            }
        }
        /// <summary>
        /// Allow as page field?
        /// </summary>
        internal bool DragToPage
        {
            get
            {
                return this.GetXmlNodeBool("@dragToPage", true);
            }
        }
    }
}
