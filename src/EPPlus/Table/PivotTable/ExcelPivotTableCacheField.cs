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

using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Core;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// A pivot tables cache field
/// </summary>
public class ExcelPivotTableCacheField : XmlHelper
{
    [Flags]
    private enum DataTypeFlags
    {
        Empty = 0x1,
        String = 0x2,
        Int = 0x4,
        Number = 0x8,
        DateTime = 0x10,
        Boolean = 0x20,
        Error = 0x30,
        Float = 0x40,
    }

    internal PivotTableCacheInternal _cache;

    internal ExcelPivotTableCacheField(XmlNamespaceManager nsm, XmlNode topNode, PivotTableCacheInternal cache, int index)
        : base(nsm, topNode)
    {
        this._cache = cache;
        this.Index = index;
        this.SetCacheFieldNode();

        if (this.NumFmtId.HasValue)
        {
            ExcelStyles? styles = cache._wb.Styles;
            int ix = styles.NumberFormats.FindIndexById(this.NumFmtId.Value.ToString(CultureInfo.InvariantCulture));

            if (ix >= 0)
            {
                this.Format = styles.NumberFormats[ix].Format;
            }
        }
    }

    /// <summary>
    /// The index in the collection of the pivot field
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// The name for the field
    /// </summary>
    public string Name
    {
        get { return this.GetXmlNodeString("@name"); }
        internal set { this.SetXmlNodeString("@name", value); }
    }

    /// <summary>
    /// A list of unique items for the field 
    /// </summary>
    public EPPlusReadOnlyList<object> SharedItems { get; } = new EPPlusReadOnlyList<object>();

    /// <summary>
    /// A list of group items, if the field has grouping.
    /// <seealso cref="Grouping"/>
    /// </summary>
    public EPPlusReadOnlyList<object> GroupItems { get; set; } = new EPPlusReadOnlyList<object>();

    internal Dictionary<object, int> _cacheLookup = null;

    /// <summary>
    /// The type of date grouping
    /// </summary>
    public eDateGroupBy DateGrouping { get; private set; }

    /// <summary>
    /// Grouping proprerties, if the field has grouping
    /// </summary>
    public ExcelPivotTableFieldGroup Grouping { get; set; }

    /// <summary>
    /// The number format for the field
    /// </summary>
    public string Format { get; set; }

    internal int? NumFmtId
    {
        get { return this.GetXmlNodeIntNull("@numFmtId"); }
        set { this.SetXmlNodeInt("@numFmtId", value); }
    }

    internal void WriteSharedItems(XmlElement fieldNode, XmlNamespaceManager nsm)
    {
        XmlElement? shNode = (XmlElement)fieldNode.SelectSingleNode("d:sharedItems", nsm);
        shNode.RemoveAll();

        DataTypeFlags flags = this.GetFlags();

        this._cacheLookup = new Dictionary<object, int>(new CacheComparer());

        if (this.IsRowColumnOrPage || this.HasSlicer)
        {
            this.AppendSharedItems(shNode);
        }

        int noTypes = GetNoOfTypes(flags);

        if (noTypes > 1
            && flags != (DataTypeFlags.Int | DataTypeFlags.Number)
            && flags != (DataTypeFlags.Float | DataTypeFlags.Number)
            && flags != (DataTypeFlags.Int | DataTypeFlags.Float | DataTypeFlags.Number)
            && flags != (DataTypeFlags.Int | DataTypeFlags.Number | DataTypeFlags.Empty)
            && flags != (DataTypeFlags.Float | DataTypeFlags.Number | DataTypeFlags.Empty)
            && flags != (DataTypeFlags.Int | DataTypeFlags.Float | DataTypeFlags.Number | DataTypeFlags.Empty)
            && this.SharedItems.Count > 1)
        {
            if ((flags & DataTypeFlags.String) == DataTypeFlags.String || (flags & DataTypeFlags.String) == DataTypeFlags.Empty)
            {
                shNode.SetAttribute("containsMixedTypes", "1");
            }
            else
            {
                shNode.SetAttribute("containsMixedTypes", "1");
            }

            SetFlags(shNode, flags);
        }
        else
        {
            if ((flags & DataTypeFlags.String) != DataTypeFlags.String
                && (flags & DataTypeFlags.Empty) != DataTypeFlags.Empty
                && (flags & DataTypeFlags.Boolean) != DataTypeFlags.Boolean)
            {
                shNode.SetAttribute("containsSemiMixedTypes", "0");
                shNode.SetAttribute("containsString", "0");
            }

            SetFlags(shNode, flags);
        }
    }

    internal bool IsRowColumnOrPage
    {
        get
        {
            foreach (ExcelPivotTable? pt in this._cache._pivotTables)
            {
                if (this.Index < pt.Fields.Count)
                {
                    ePivotFieldAxis axis = pt.Fields[this.Index].Axis;

                    if (axis == ePivotFieldAxis.Column || axis == ePivotFieldAxis.Row || axis == ePivotFieldAxis.Page)
                    {
                        return true;
                    }
                }
                else
                {
                    return false;
                }
            }

            return false;
        }
    }

    /// <summary>
    /// The formula for cache field.
    /// The formula for the calculated field. 
    /// Note: In formulas you create for calculated fields or calculated items, you can use operators and expressions as you do in other worksheet formulas. 
    /// You can use constants and refer to data from the pivot table, but you cannot use cell references or defined names.You cannot use worksheet functions that require cell references or defined names as arguments, and you cannot use array functions.
    /// <seealso cref="ExcelPivotTableFieldCollection.AddCalculatedField(string, string)"/>
    /// </summary>
    public string Formula
    {
        get { return this.GetXmlNodeString("@formula"); }
        set
        {
            if (this.DatabaseField)
            {
                throw new InvalidOperationException("Can't set a formula to a database field");
            }

            if (string.IsNullOrEmpty(value) || value.Trim() == "")
            {
                throw new ArgumentException("The formula can't be blank", "formula");
            }

            this.SetXmlNodeString("@formula", value);
        }
    }

    internal bool DatabaseField
    {
        get { return this.GetXmlNodeBool("@databaseField", true); }
        set { this.SetXmlNodeBool("@databaseField", value, true); }
    }

    internal bool HasSlicer
    {
        get
        {
            foreach (ExcelPivotTable? pt in this._cache._pivotTables)
            {
                if (pt.Fields.Count > this.Index && pt.Fields[this.Index].Slicer != null)
                {
                    return true;
                }
            }

            return false;
        }
    }

    internal void UpdateSlicers()
    {
        foreach (ExcelPivotTable? pt in this._cache._pivotTables)
        {
            ExcelPivotTableSlicer? s = pt.Fields[this.Index].Slicer;

            if (s != null)
            {
                s.Cache.Data.Items.RefreshMe();
            }
        }
    }

    private static void SetFlags(XmlElement shNode, DataTypeFlags flags)
    {
        if ((flags & DataTypeFlags.DateTime) == DataTypeFlags.DateTime)
        {
            shNode.SetAttribute("containsDate", "1");
        }

        if ((flags & DataTypeFlags.Number) == DataTypeFlags.Number)
        {
            shNode.SetAttribute("containsNumber", "1");
        }

        if ((flags & DataTypeFlags.Int) == DataTypeFlags.Int && (flags & DataTypeFlags.Float) != DataTypeFlags.Float)
        {
            shNode.SetAttribute("containsInteger", "1");
        }

        if ((flags & DataTypeFlags.Empty) == DataTypeFlags.Empty)
        {
            shNode.SetAttribute("containsBlank", "1");
        }

        if ((flags & DataTypeFlags.String) != DataTypeFlags.String
            && (flags & DataTypeFlags.Boolean) != DataTypeFlags.Boolean
            && (flags & DataTypeFlags.Error) != DataTypeFlags.Error)
        {
            shNode.SetAttribute("containsString", "0");
        }
    }

    private static int GetNoOfTypes(DataTypeFlags flags)
    {
        int types = 0;

        foreach (DataTypeFlags v in Enum.GetValues(typeof(DataTypeFlags)))
        {
            if (v != DataTypeFlags.Empty && (flags & v) == v)
            {
                types++;
            }
        }

        return types;
    }

    private void AppendSharedItems(XmlElement shNode)
    {
        int index = 0;
        bool isLongText = false;

        foreach (object? si in this.SharedItems)
        {
            if (si == null || si.Equals(ExcelPivotTable.PivotNullValue))
            {
                this._cacheLookup.Add(ExcelPivotTable.PivotNullValue, index++);
                AppendItem(shNode, "m", null);
            }
            else
            {
                this._cacheLookup.Add(si, index++);
                Type? t = si.GetType();
                TypeCode tc = Type.GetTypeCode(t);

                switch (tc)
                {
                    case TypeCode.Byte:
                    case TypeCode.SByte:
                    case TypeCode.UInt16:
                    case TypeCode.UInt32:
                    case TypeCode.UInt64:
                    case TypeCode.Int16:
                    case TypeCode.Int32:
                    case TypeCode.Int64:
                    case TypeCode.Decimal:
                    case TypeCode.Double:
                    case TypeCode.Single:
                        if (t.IsEnum)
                        {
                            AppendItem(shNode, "s", si.ToString());
                        }
                        else
                        {
                            AppendItem(shNode, "n", ConvertUtil.GetValueForXml(si, false));
                        }

                        break;

                    case TypeCode.DateTime:
                        DateTime d = (DateTime)si;

                        if (d.Year > 1899)
                        {
                            AppendItem(shNode, "d", d.ToString("s"));
                        }
                        else
                        {
                            AppendItem(shNode, "d", d.ToString("HH:mm:ss", CultureInfo.InvariantCulture));
                        }

                        break;

                    case TypeCode.Boolean:
                        AppendItem(shNode, "b", ConvertUtil.GetValueForXml(si, false));

                        break;

                    case TypeCode.Empty:
                        AppendItem(shNode, "m", null);

                        break;

                    default:
                        if (t == typeof(TimeSpan))
                        {
                            d = new DateTime(((TimeSpan)si).Ticks);

                            if (d.Year > 1899)
                            {
                                AppendItem(shNode, "d", d.ToString("s"));
                            }
                            else
                            {
                                AppendItem(shNode, "d", d.ToString("HH:mm:ss", CultureInfo.InvariantCulture));
                            }
                        }
                        else if (t == typeof(ExcelErrorValue))
                        {
                            AppendItem(shNode, "e", si.ToString());
                        }
                        else
                        {
                            string? s = si.ToString();
                            AppendItem(shNode, "s", s);

                            if (s.Length > 255 && isLongText == false)
                            {
                                isLongText = true;
                            }
                        }

                        break;
                }
            }
        }

        if (isLongText)
        {
            shNode.SetAttribute("longText", "1");
        }
    }

    private DataTypeFlags GetFlags()
    {
        DataTypeFlags flags = 0;

        foreach (object? si in this.SharedItems)
        {
            if (si == null || si.Equals(ExcelPivotTable.PivotNullValue))
            {
                flags |= DataTypeFlags.Empty;
            }
            else
            {
                Type? t = si.GetType();

                switch (Type.GetTypeCode(t))
                {
                    case TypeCode.String:
                    case TypeCode.Char:
                        flags |= DataTypeFlags.String;

                        break;

                    case TypeCode.Byte:
                    case TypeCode.SByte:
                    case TypeCode.UInt16:
                    case TypeCode.UInt32:
                    case TypeCode.UInt64:
                    case TypeCode.Int16:
                    case TypeCode.Int32:
                    case TypeCode.Int64:
                        if (t.IsEnum)
                        {
                            flags |= DataTypeFlags.String;
                        }
                        else
                        {
                            flags |= DataTypeFlags.Number | DataTypeFlags.Int;
                        }

                        break;

                    case TypeCode.Decimal:
                    case TypeCode.Double:
                    case TypeCode.Single:
                        flags |= DataTypeFlags.Number;

                        if ((flags & DataTypeFlags.Int) != DataTypeFlags.Int && Convert.ToDouble(si) % 1 == 0)
                        {
                            flags |= DataTypeFlags.Int;
                        }
                        else if ((flags & DataTypeFlags.Float) != DataTypeFlags.Float && Convert.ToDouble(si) % 1 != 0)
                        {
                            flags |= DataTypeFlags.Float;
                        }

                        break;

                    case TypeCode.DateTime:
                        flags |= DataTypeFlags.DateTime;

                        break;

                    case TypeCode.Boolean:
                        flags |= DataTypeFlags.Boolean;

                        break;

                    case TypeCode.Empty:
                        flags |= DataTypeFlags.Empty;

                        break;

                    default:
                        if (t == typeof(TimeSpan))
                        {
                            flags |= DataTypeFlags.DateTime;
                        }
                        else if (t == typeof(ExcelErrorValue))
                        {
                            flags |= DataTypeFlags.Error;
                        }
                        else
                        {
                            flags |= DataTypeFlags.String;
                        }

                        break;
                }
            }
        }

        return flags;
    }

    private static void AppendItem(XmlElement shNode, string elementName, string value)
    {
        XmlElement? e = shNode.OwnerDocument.CreateElement(elementName, ExcelPackage.schemaMain);

        if (value != null)
        {
            e.SetAttribute("v", value);
        }

        shNode.AppendChild(e);
    }

    internal void SetCacheFieldNode()
    {
        XmlNode? groupNode = this.GetNode("d:fieldGroup");

        if (groupNode != null)
        {
            XmlNode? groupBy = groupNode.SelectSingleNode("d:rangePr/@groupBy", this.NameSpaceManager);

            if (groupBy == null)
            {
                this.Grouping = new ExcelPivotTableFieldNumericGroup(this.NameSpaceManager, this.TopNode);
            }
            else
            {
                this.DateGrouping = (eDateGroupBy)Enum.Parse(typeof(eDateGroupBy), groupBy.Value, true);
                this.Grouping = new ExcelPivotTableFieldDateGroup(this.NameSpaceManager, groupNode);
            }

            XmlNode? groupItems = groupNode.SelectSingleNode("d:groupItems", this.NameSpaceManager);

            if (groupItems != null)
            {
                this.AddItems(this.GroupItems, groupItems, true);
            }
        }

        XmlNode? si = this.GetNode("d:sharedItems");

        if (si != null)
        {
            this.AddItems(this.SharedItems, si, groupNode == null);
        }
    }

    private void AddItems(EPPlusReadOnlyList<Object> items, XmlNode itemsNode, bool updateCacheLookup)
    {
        if (updateCacheLookup)
        {
            this._cacheLookup = new Dictionary<object, int>(new CacheComparer());
        }

        foreach (XmlElement c in itemsNode.ChildNodes)
        {
            if (c.LocalName == "s")
            {
                items.Add(c.Attributes["v"].Value);
            }
            else if (c.LocalName == "d")
            {
                if (ConvertUtil.TryParseDateString(c.Attributes["v"].Value, out DateTime d))
                {
                    items.Add(d);
                }
                else
                {
                    items.Add(c.Attributes["v"].Value);
                }
            }
            else if (c.LocalName == "n")
            {
                if (ConvertUtil.TryParseNumericString(c.Attributes["v"].Value, out double num))
                {
                    items.Add(num);
                }
                else
                {
                    items.Add(c.Attributes["v"].Value);
                }
            }
            else if (c.LocalName == "b")
            {
                if (ConvertUtil.TryParseBooleanString(c.Attributes["v"].Value, out bool b))
                {
                    items.Add(b);
                }
                else
                {
                    items.Add(c.Attributes["v"].Value);
                }
            }
            else if (c.LocalName == "e")
            {
                if (ExcelErrorValue.Values.StringIsErrorValue(c.Attributes["v"].Value))
                {
                    items.Add(ExcelErrorValue.Parse(c.Attributes["v"].Value));
                }
                else
                {
                    items.Add(c.Attributes["v"].Value);
                }
            }
            else
            {
                items.Add(ExcelPivotTable.PivotNullValue);
            }

            if (updateCacheLookup)
            {
                object? key = items[items.Count - 1];

                if (this._cacheLookup.ContainsKey(key))
                {
                    items._list.Remove(key);
                }
                else
                {
                    this._cacheLookup.Add(key, items.Count - 1);
                }
            }
        }
    }

    #region Grouping

    internal ExcelPivotTableFieldDateGroup SetDateGroup(ExcelPivotTableField field, eDateGroupBy groupBy, DateTime StartDate, DateTime EndDate, int interval)
    {
        ExcelPivotTableFieldDateGroup group = new(this.NameSpaceManager, this.TopNode);
        this.SetXmlNodeBool("d:sharedItems/@containsDate", true);
        this.SetXmlNodeBool("d:sharedItems/@containsNonDate", false);
        this.SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);

        group.TopNode.InnerXml += string.Format("<fieldGroup base=\"{0}\"><rangePr groupBy=\"{1}\" /><groupItems /></fieldGroup>",
                                                field.BaseIndex,
                                                groupBy.ToString().ToLower(CultureInfo.InvariantCulture));

        if (StartDate.Year < 1900)
        {
            this.SetXmlNodeString("d:fieldGroup/d:rangePr/@startDate", "1900-01-01T00:00:00");
        }
        else
        {
            this.SetXmlNodeString("d:fieldGroup/d:rangePr/@startDate", StartDate.ToString("s", CultureInfo.InvariantCulture));
            this.SetXmlNodeString("d:fieldGroup/d:rangePr/@autoStart", "0");
        }

        if (EndDate == DateTime.MaxValue)
        {
            this.SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", "9999-12-31T00:00:00");
        }
        else
        {
            this.SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", EndDate.ToString("s", CultureInfo.InvariantCulture));
            this.SetXmlNodeString("d:fieldGroup/d:rangePr/@autoEnd", "0");
        }

        int items = this.AddDateGroupItems(group, groupBy, StartDate, EndDate, interval);

        this.Grouping = group;
        this.DateGrouping = groupBy;

        return group;
    }

    internal ExcelPivotTableFieldNumericGroup SetNumericGroup(int baseIndex, double start, double end, double interval)
    {
        ExcelPivotTableFieldNumericGroup group = new(this.NameSpaceManager, this.TopNode);
        this.SetXmlNodeBool("d:sharedItems/@containsNumber", true);
        this.SetXmlNodeBool("d:sharedItems/@containsInteger", true);
        this.SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);
        this.SetXmlNodeBool("d:sharedItems/@containsString", false);

        group.TopNode.InnerXml +=
            string.Format("<fieldGroup base=\"{0}\"><rangePr autoStart=\"0\" autoEnd=\"0\" startNum=\"{1}\" endNum=\"{2}\" groupInterval=\"{3}\"/><groupItems /></fieldGroup>",
                          baseIndex,
                          start.ToString(CultureInfo.InvariantCulture),
                          end.ToString(CultureInfo.InvariantCulture),
                          interval.ToString(CultureInfo.InvariantCulture));

        int items = this.AddNumericGroupItems(group, start, end, interval);
        this.Grouping = group;

        return group;
    }

    private int AddNumericGroupItems(ExcelPivotTableFieldNumericGroup group, double start, double end, double interval)
    {
        if (interval < 0)
        {
            throw new Exception("The interval must be a positiv");
        }

        if (start > end)
        {
            throw new Exception("Then End number must be larger than the Start number");
        }

        XmlElement groupItemsNode = group.TopNode.SelectSingleNode("d:fieldGroup/d:groupItems", group.NameSpaceManager) as XmlElement;
        int items = 2;

        //First date
        double index = start;
        double nextIndex = start + interval;
        this.GroupItems.Clear();
        this.AddGroupItem(groupItemsNode, "<" + start.ToString(CultureInfo.CurrentCulture));

        while (index < end)
        {
            this.AddGroupItem(groupItemsNode,
                              string.Format("{0}-{1}", index.ToString(CultureInfo.CurrentCulture), nextIndex.ToString(CultureInfo.CurrentCulture)));

            index = nextIndex;
            nextIndex += interval;
            items++;
        }

        this.AddGroupItem(groupItemsNode, ">" + index.ToString(CultureInfo.CurrentCulture));

        this.UpdateCacheLookupFromItems(this.GroupItems._list);

        return items;
    }

    private int AddDateGroupItems(ExcelPivotTableFieldGroup group, eDateGroupBy GroupBy, DateTime StartDate, DateTime EndDate, int interval)
    {
        XmlElement groupItemsNode = group.TopNode.SelectSingleNode("d:fieldGroup/d:groupItems", group.NameSpaceManager) as XmlElement;
        int items = 2;
        this.GroupItems.Clear();

        //First date
        this.AddGroupItem(groupItemsNode, "<" + StartDate.ToString("s", CultureInfo.InvariantCulture).Substring(0, 10));

        switch (GroupBy)
        {
            case eDateGroupBy.Seconds:
            case eDateGroupBy.Minutes:
                this.AddTimeSerie(60, groupItemsNode);
                items += 60;

                break;

            case eDateGroupBy.Hours:
                this.AddTimeSerie(24, groupItemsNode);
                items += 24;

                break;

            case eDateGroupBy.Days:
                if (interval == 1)
                {
                    DateTime dt = new DateTime(2008, 1, 1); //pick a year with 366 days

                    while (dt.Year == 2008)
                    {
                        this.AddGroupItem(groupItemsNode, dt.ToString("dd-MMM"));
                        dt = dt.AddDays(1);
                    }

                    items += 366;
                }
                else
                {
                    DateTime dt = StartDate;
                    items = 0;

                    while (dt < EndDate)
                    {
                        this.AddGroupItem(groupItemsNode, dt.ToString("dd-MMM"));
                        dt = dt.AddDays(interval);
                        items++;
                    }
                }

                break;

            case eDateGroupBy.Months:
                this.AddGroupItem(groupItemsNode, "jan");
                this.AddGroupItem(groupItemsNode, "feb");
                this.AddGroupItem(groupItemsNode, "mar");
                this.AddGroupItem(groupItemsNode, "apr");
                this.AddGroupItem(groupItemsNode, "may");
                this.AddGroupItem(groupItemsNode, "jun");
                this.AddGroupItem(groupItemsNode, "jul");
                this.AddGroupItem(groupItemsNode, "aug");
                this.AddGroupItem(groupItemsNode, "sep");
                this.AddGroupItem(groupItemsNode, "oct");
                this.AddGroupItem(groupItemsNode, "nov");
                this.AddGroupItem(groupItemsNode, "dec");
                items += 12;

                break;

            case eDateGroupBy.Quarters:
                this.AddGroupItem(groupItemsNode, "Qtr1");
                this.AddGroupItem(groupItemsNode, "Qtr2");
                this.AddGroupItem(groupItemsNode, "Qtr3");
                this.AddGroupItem(groupItemsNode, "Qtr4");
                items += 4;

                break;

            case eDateGroupBy.Years:
                if (StartDate.Year >= 1900 && EndDate != DateTime.MaxValue)
                {
                    for (int year = StartDate.Year; year <= EndDate.Year; year++)
                    {
                        this.AddGroupItem(groupItemsNode, year.ToString());
                    }

                    items += EndDate.Year - StartDate.Year + 1;
                }

                break;

            default:
                throw new Exception("unsupported grouping");
        }

        //Lastdate
        this.AddGroupItem(groupItemsNode, ">" + EndDate.ToString("s", CultureInfo.InvariantCulture).Substring(0, 10));

        this.UpdateCacheLookupFromItems(this.GroupItems._list);

        return items;
    }

    private void UpdateCacheLookupFromItems(List<object> items)
    {
        this._cacheLookup = new Dictionary<object, int>(new CacheComparer());

        for (int i = 0; i < items.Count; i++)
        {
            object? key = items[i];

            if (!this._cacheLookup.ContainsKey(key))
            {
                this._cacheLookup.Add(key, i);
            }
        }
    }

    private void AddTimeSerie(int count, XmlElement groupItems)
    {
        for (int i = 0; i < count; i++)
        {
            this.AddGroupItem(groupItems, string.Format("{0:00}", i));
        }
    }

    private void AddGroupItem(XmlElement groupItems, string value)
    {
        XmlElement? s = groupItems.OwnerDocument.CreateElement("s", ExcelPackage.schemaMain);
        s.SetAttribute("v", value);
        groupItems.AppendChild(s);
        this.GroupItems.Add(value);
    }

    #endregion

    internal void Refresh()
    {
        if (!string.IsNullOrEmpty(this.Formula))
        {
            return;
        }

        if (this.Grouping == null)
        {
            this.UpdateSharedItems();
        }
        else
        {
            this.UpdateGroupItems();
        }
    }

    private void UpdateGroupItems()
    {
        foreach (ExcelPivotTable? pt in this._cache._pivotTables)
        {
            if (pt.Fields[this.Index].IsRowField
                || pt.Fields[this.Index].IsColumnField
                || pt.Fields[this.Index].IsPageField
                || pt.Fields[this.Index].Cache.HasSlicer)
            {
                if (pt.Fields[this.Index].Items.Count == 0)
                {
                    pt.Fields[this.Index].UpdateGroupItems(this, true);
                }
            }
            else
            {
                pt.Fields[this.Index].DeleteNode("d:items");
            }
        }
    }

    private void UpdateSharedItems()
    {
        ExcelRangeBase? range = this._cache.SourceRange;

        if (range == null)
        {
            return;
        }

        int column = range._fromCol + this.Index;
        HashSet<object>? hs = new HashSet<object>(new InvariantObjectComparer());
        ExcelWorksheet? ws = range.Worksheet;
        int dimensionToRow = ws.Dimension?._toRow ?? range._fromRow + 1;
        int toRow = range._toRow < dimensionToRow ? range._toRow : dimensionToRow;

        //Get unique values.
        for (int row = range._fromRow + 1; row <= toRow; row++)
        {
            AddSharedItemToHashSet(hs, ws.GetValue(row, column));
        }

        //A pivot table cache can reference multiple Pivot tables, so we need to update them all
        foreach (ExcelPivotTable? pt in this._cache._pivotTables)
        {
            HashSet<object>? existingItems = new HashSet<object>();
            List<ExcelPivotTableFieldItem>? list = pt.Fields[this.Index].Items._list;

            for (int ix = 0; ix < list.Count; ix++)
            {
                object? v = list[ix].Value ?? ExcelPivotTable.PivotNullValue;

                if (!hs.Contains(v) || existingItems.Contains(v))
                {
                    list.RemoveAt(ix);
                    ix--;
                }
                else
                {
                    existingItems.Add(v);
                }
            }

            int hasSubTotalSubt = list.Count > 0 && list[list.Count - 1].Type == eItemType.Default ? 1 : 0;

            foreach (object? c in hs)
            {
                if (!existingItems.Contains(c))
                {
                    list.Insert(list.Count - hasSubTotalSubt, new ExcelPivotTableFieldItem() { Value = c });
                }
            }

            if (list.Count > 0 && list[list.Count - 1].Type != eItemType.Default && pt.Fields[this.Index].GetXmlNodeBool("@defaultSubtotal", true) == true)
            {
                list.Add(new ExcelPivotTableFieldItem() { Type = eItemType.Default, X = -1 });
            }
        }

        this.SharedItems._list = hs.ToList();
        this.UpdateCacheLookupFromItems(this.SharedItems._list);

        if (this.HasSlicer)
        {
            this.UpdateSlicers();
        }
    }

    internal static object AddSharedItemToHashSet(HashSet<object> hs, object o)
    {
        if (o == null)
        {
            o = ExcelPivotTable.PivotNullValue;
        }
        else
        {
            Type? t = o.GetType();

            if (t == typeof(TimeSpan))
            {
                long ticks = ((TimeSpan)o).Ticks + (TimeSpan.TicksPerSecond / 2);
                o = new DateTime(ticks - (ticks % TimeSpan.TicksPerSecond));
            }

            if (t == typeof(DateTime))
            {
                long ticks = ((DateTime)o).Ticks;

                if (ticks % TimeSpan.TicksPerSecond != 0)
                {
                    ticks += TimeSpan.TicksPerSecond / 2;
                    o = new DateTime(ticks - (ticks % TimeSpan.TicksPerSecond));
                }
            }
        }

        if (!hs.Contains(o))
        {
            hs.Add(o);
        }

        return o;
    }
}

internal class CacheComparer : IEqualityComparer<object>
{
    public new bool Equals(object x, object y)
    {
        x = GetCaseInsensitiveValue(x);
        y = GetCaseInsensitiveValue(y);

        return x.Equals(y);
    }

    private static object GetCaseInsensitiveValue(object x)
    {
        if (x is string sx)
        {
            x = sx.ToLower();
        }
        else if (x is char cx)
        {
            x = char.ToLower(cx);
        }

        return x;
    }

    public int GetHashCode(object obj)
    {
        return GetCaseInsensitiveValue(obj).GetHashCode();
    }
}