/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/02/2020         EPPlus Software AB       EPPlus 5.4
 *************************************************************************************************/
using OfficeOpenXml;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable.Filter;

/// <summary>
/// Defines a pivot table filter
/// </summary>
public class ExcelPivotTableFilter : XmlHelper
{
    XmlNode _filterColumnNode;
    bool _date1904;
    internal ExcelPivotTableFilter(XmlNamespaceManager nsm, XmlNode topNode, bool date1904) : base(nsm, topNode)
    {
        if (topNode.InnerXml == "")
        {
            topNode.InnerXml = "<autoFilter ref=\"A1\"><filterColumn colId=\"0\"></filterColumn></autoFilter>";
        }

        this._filterColumnNode = this.GetNode("d:autoFilter/d:filterColumn");
        this._date1904 = date1904;
    } 
    /// <summary>
    /// The id 
    /// </summary>
    public int Id
    {
        get
        {
            return this.GetXmlNodeInt("@id");
        }
        internal set
        {
            this.SetXmlNodeInt("@id", value);
        }
    }
    /// <summary>
    /// The name of the pivot filter
    /// </summary>
    public string Name
    {
        get
        {
            return this.GetXmlNodeString("@name");
        }
        set
        {
            this.SetXmlNodeString("@name", value, true);
        }
    }
    /// <summary>
    /// The description of the pivot filter
    /// </summary>
    public string Description
    {
        get
        {
            return this.GetXmlNodeString("@description");
        }
        set
        {
            this.SetXmlNodeString("@description", value, true);
        }
    }
    internal void CreateDateCustomFilter(ePivotTableDateValueFilterType type)
    {
        this._filterColumnNode.InnerXml = "<customFilters/>";
        ExcelCustomFilterColumn? cf = new ExcelCustomFilterColumn(this.NameSpaceManager, this._filterColumnNode);

        eFilterOperator t;
        string? v = ConvertUtil.GetValueForXml(this.Value1, this._date1904);
        switch (type)
        {
            case ePivotTableDateValueFilterType.DateNotEqual:
                t = eFilterOperator.NotEqual;
                break;
            case ePivotTableDateValueFilterType.DateNewerThan:
            case ePivotTableDateValueFilterType.DateBetween:
                t = eFilterOperator.GreaterThan;
                break;
            case ePivotTableDateValueFilterType.DateNewerThanOrEqual:
                t = eFilterOperator.GreaterThanOrEqual;
                break;
            case ePivotTableDateValueFilterType.DateOlderThan:
            case ePivotTableDateValueFilterType.DateNotBetween:
                t = eFilterOperator.LessThan;
                break;
            case ePivotTableDateValueFilterType.DateOlderThanOrEqual:
                t = eFilterOperator.LessThanOrEqual;
                break;
            default:
                t = eFilterOperator.Equal;
                break;
        }

        ExcelFilterCustomItem? item1 = new ExcelFilterCustomItem(v, t);
        cf.Filters.Add(item1);

        if (type == ePivotTableDateValueFilterType.DateBetween)
        {
            cf.And = true;
            cf.Filters.Add(new ExcelFilterCustomItem(ConvertUtil.GetValueForXml(this.Value2, this._date1904), eFilterOperator.LessThanOrEqual));
        }
        else if (type == ePivotTableDateValueFilterType.DateNotBetween)
        {
            cf.And = false;
            cf.Filters.Add(new ExcelFilterCustomItem(ConvertUtil.GetValueForXml(this.Value2, this._date1904), eFilterOperator.GreaterThan));
        }

        this._filter = cf;
    }

    internal void CreateDateDynamicFilter(ePivotTableDatePeriodFilterType type)
    {
        this._filterColumnNode.InnerXml = "<dynamicFilter />";
        ExcelDynamicFilterColumn? df = new ExcelDynamicFilterColumn(this.NameSpaceManager, this._filterColumnNode);
        switch(type)
        {
            case ePivotTableDatePeriodFilterType.LastMonth:
                df.Type = eDynamicFilterType.LastMonth;
                break;
            case ePivotTableDatePeriodFilterType.LastQuarter:
                df.Type = eDynamicFilterType.LastQuarter;
                break;
            case ePivotTableDatePeriodFilterType.LastWeek:
                df.Type = eDynamicFilterType.LastWeek;
                break;
            case ePivotTableDatePeriodFilterType.LastYear:
                df.Type = eDynamicFilterType.LastYear;
                break;
            case ePivotTableDatePeriodFilterType.M1:
                df.Type = eDynamicFilterType.M1;
                break;
            case ePivotTableDatePeriodFilterType.M2:
                df.Type = eDynamicFilterType.M2;
                break;
            case ePivotTableDatePeriodFilterType.M3:
                df.Type = eDynamicFilterType.M3;
                break;
            case ePivotTableDatePeriodFilterType.M4:
                df.Type = eDynamicFilterType.M4;
                break;
            case ePivotTableDatePeriodFilterType.M5:
                df.Type = eDynamicFilterType.M5;
                break;
            case ePivotTableDatePeriodFilterType.M6:
                df.Type = eDynamicFilterType.M6;
                break;
            case ePivotTableDatePeriodFilterType.M7:
                df.Type = eDynamicFilterType.M7;
                break;
            case ePivotTableDatePeriodFilterType.M8:
                df.Type = eDynamicFilterType.M8;
                break;
            case ePivotTableDatePeriodFilterType.M9:
                df.Type = eDynamicFilterType.M9;
                break;
            case ePivotTableDatePeriodFilterType.M10:
                df.Type = eDynamicFilterType.M10;
                break;
            case ePivotTableDatePeriodFilterType.M11:
                df.Type = eDynamicFilterType.M11;
                break;
            case ePivotTableDatePeriodFilterType.M12:
                df.Type = eDynamicFilterType.M12;
                break;
            case ePivotTableDatePeriodFilterType.NextMonth:
                df.Type = eDynamicFilterType.NextMonth;
                break;
            case ePivotTableDatePeriodFilterType.NextQuarter:
                df.Type = eDynamicFilterType.NextQuarter;
                break;
            case ePivotTableDatePeriodFilterType.NextWeek:
                df.Type = eDynamicFilterType.NextWeek;
                break;
            case ePivotTableDatePeriodFilterType.NextYear:
                df.Type = eDynamicFilterType.NextYear;
                break;
            case ePivotTableDatePeriodFilterType.Q1:
                df.Type = eDynamicFilterType.Q1;
                break;
            case ePivotTableDatePeriodFilterType.Q2:
                df.Type = eDynamicFilterType.Q2;
                break;
            case ePivotTableDatePeriodFilterType.Q3:
                df.Type = eDynamicFilterType.Q3;
                break;
            case ePivotTableDatePeriodFilterType.Q4:
                df.Type = eDynamicFilterType.Q4;
                break;
            case ePivotTableDatePeriodFilterType.ThisMonth:
                df.Type = eDynamicFilterType.ThisMonth;
                break;
            case ePivotTableDatePeriodFilterType.ThisQuarter:
                df.Type = eDynamicFilterType.ThisQuarter;
                break;
            case ePivotTableDatePeriodFilterType.ThisWeek:
                df.Type = eDynamicFilterType.ThisWeek;
                break;
            case ePivotTableDatePeriodFilterType.ThisYear:
                df.Type = eDynamicFilterType.ThisYear;
                break;
            case ePivotTableDatePeriodFilterType.Yesterday:
                df.Type = eDynamicFilterType.Yesterday;
                break;
            case ePivotTableDatePeriodFilterType.Today:
                df.Type = eDynamicFilterType.Today;
                break;
            case ePivotTableDatePeriodFilterType.Tomorrow:
                df.Type = eDynamicFilterType.Tomorrow;
                break;
            case ePivotTableDatePeriodFilterType.YearToDate:
                df.Type = eDynamicFilterType.YearToDate;
                break;
            default:
                throw new Exception($"Unsupported Pivottable filter type {type}");
        }

        this._filter = df;
    }

    internal void CreateTop10Filter(ePivotTableTop10FilterType type, bool isTop, double value)
    {
        this._filterColumnNode.InnerXml = "<top10 />";
        ExcelTop10FilterColumn? tf = new ExcelTop10FilterColumn(this.NameSpaceManager, this._filterColumnNode);

        tf.Percent = (type == ePivotTableTop10FilterType.Percent);
        tf.Top = isTop;
        tf.Value = value;
        tf.FilterValue = value;

        this._filter = tf;
    }

    internal void CreateCaptionCustomFilter(ePivotTableCaptionFilterType type)
    {
        this._filterColumnNode.InnerXml = "<customFilters/>";
        ExcelCustomFilterColumn? cf = new ExcelCustomFilterColumn(this.NameSpaceManager, this._filterColumnNode);

        eFilterOperator t;
        string? v = this.StringValue1;
        switch(type)
        {
            case ePivotTableCaptionFilterType.CaptionNotBeginsWith:
            case ePivotTableCaptionFilterType.CaptionNotContains:
            case ePivotTableCaptionFilterType.CaptionNotEndsWith:
            case ePivotTableCaptionFilterType.CaptionNotEqual:
                t = eFilterOperator.NotEqual;
                break;
            case ePivotTableCaptionFilterType.CaptionGreaterThan:
                t = eFilterOperator.GreaterThan;
                break;
            case ePivotTableCaptionFilterType.CaptionGreaterThanOrEqual:
            case ePivotTableCaptionFilterType.CaptionBetween:
                t = eFilterOperator.GreaterThanOrEqual;
                break;
            case ePivotTableCaptionFilterType.CaptionLessThan:
            case ePivotTableCaptionFilterType.CaptionNotBetween:
                t = eFilterOperator.LessThan;
                break;
            case ePivotTableCaptionFilterType.CaptionLessThanOrEqual:
                t = eFilterOperator.LessThanOrEqual;
                break;
            default:
                t = eFilterOperator.Equal;
                break;
        }
        switch (type)
        {
            case ePivotTableCaptionFilterType.CaptionBeginsWith:
            case ePivotTableCaptionFilterType.CaptionNotBeginsWith:
                v += "*";
                break;
            case ePivotTableCaptionFilterType.CaptionContains:
            case ePivotTableCaptionFilterType.CaptionNotContains:
                v = $"*{v}*";
                break;
            case ePivotTableCaptionFilterType.CaptionEndsWith:
            case ePivotTableCaptionFilterType.CaptionNotEndsWith:
                v = $"*{v}";
                break;
        }
        ExcelFilterCustomItem? item1 = new ExcelFilterCustomItem(v, t);
        cf.Filters.Add(item1);

        if(type==ePivotTableCaptionFilterType.CaptionBetween)
        {
            cf.And = true;
            cf.Filters.Add(new ExcelFilterCustomItem(this.StringValue2, eFilterOperator.LessThanOrEqual));
        }
        else if (type == ePivotTableCaptionFilterType.CaptionNotBetween)
        {
            cf.And = false;
            cf.Filters.Add(new ExcelFilterCustomItem(this.StringValue2, eFilterOperator.GreaterThan));
        }

        this._filter = cf;
    }
    internal void CreateValueCustomFilter(ePivotTableValueFilterType type)
    {
        this._filterColumnNode.InnerXml = "<customFilters/>";
        ExcelCustomFilterColumn? cf = new ExcelCustomFilterColumn(this.NameSpaceManager, this._filterColumnNode);

        eFilterOperator t;
        string v1 = GetFilterValueAsString(this.Value1);
        switch (type)
        {
            case ePivotTableValueFilterType.ValueNotEqual:
                t = eFilterOperator.NotEqual;
                break;
            case ePivotTableValueFilterType.ValueGreaterThan:
                t = eFilterOperator.GreaterThan;
                break;
            case ePivotTableValueFilterType.ValueGreaterThanOrEqual:
            case ePivotTableValueFilterType.ValueBetween:
                t = eFilterOperator.GreaterThanOrEqual;
                break;
            case ePivotTableValueFilterType.ValueLessThan:
                t = eFilterOperator.LessThan;
                break;
            case ePivotTableValueFilterType.ValueLessThanOrEqual:
            case ePivotTableValueFilterType.ValueNotBetween:
                t = eFilterOperator.LessThanOrEqual;
                break;
            default:
                t = eFilterOperator.Equal;
                break;
        }

        ExcelFilterCustomItem? item1 = new ExcelFilterCustomItem(v1, t);
        cf.Filters.Add(item1);

        if (type == ePivotTableValueFilterType.ValueBetween)
        {
            cf.And = true;
            cf.Filters.Add(new ExcelFilterCustomItem(GetFilterValueAsString(this.Value2), eFilterOperator.LessThanOrEqual));
        }
        else if (type == ePivotTableValueFilterType.ValueNotBetween)
        {
            cf.And = false;
            cf.Filters.Add(new ExcelFilterCustomItem(GetFilterValueAsString(this.Value2), eFilterOperator.GreaterThan));
        }

        this._filter = cf;
    }

    private static string GetFilterValueAsString(object v)
    {
        if (ConvertUtil.IsNumericOrDate(v))
        {
            return ConvertUtil.GetValueDouble(v).ToString(CultureInfo.InvariantCulture);
        }
        else
        {
            return v.ToString();
        }
    }
    internal void CreateValueFilter()
    {
        this._filterColumnNode.InnerXml = "<filters/>";
        ExcelValueFilterColumn? f = new ExcelValueFilterColumn(this.NameSpaceManager, this._filterColumnNode);
        f.Filters.Add(this.StringValue1);
        this._filter = f;
    }

    /// <summary>
    /// The type of pivot filter
    /// </summary>
    public ePivotTableFilterType Type
    {
        get
        {
            return this.GetXmlNodeString("@type").ToEnum(ePivotTableFilterType.Unknown);
        }
        internal set
        {
            string? s = value.ToEnumString();
            if (s.Length <= 3 && (s[0]=='m' || s[0] == 'q'))
            {
                s = s.ToUpper();  //For M1 - M12 and Q1 - Q4
            }

            this.SetXmlNodeString("@type", s);
        }
    }
    /// <summary>
    /// The evaluation order of the pivot filter
    /// </summary>
    public int EvalOrder
    {
        get
        {
            return this.GetXmlNodeInt("@evalOrder");
        }
        internal set
        {
            this.SetXmlNodeInt("@evalOrder", value);
        }
    }
    internal int Fld
    {
        get
        {
            return this.GetXmlNodeInt("@fld");
        }
        set
        {
            this.SetXmlNodeInt("@fld", value);
        }
    }
    internal int MeasureFldIndex
    {
        get
        {
            return this.GetXmlNodeInt("@iMeasureFld");
        }
        set
        {
            this.SetXmlNodeInt("@iMeasureFld", value);
        }
    }
    internal int MeasureHierIndex
    {
        get
        {
            return this.GetXmlNodeInt("@iMeasureHier");
        }
        set
        {
            this.SetXmlNodeInt("@iMeasureHier", value);
        }
    }
    internal int MemberPropertyFldIndex
    {
        get
        {
            return this.GetXmlNodeInt("@mpFld");
        }
        set
        {
            this.SetXmlNodeInt("@mpFld", value);
        }
    }
    /// <summary>
    /// The value 1 to compare the filter to
    /// </summary>
    public object Value1
    {
        get;
        set;
    }

    /// <summary>
    /// The value 2 to compare the filter to
    /// </summary>
    public object Value2
    {
        get;
        set;
    }
    /// <summary>
    /// The string value 1 used by caption filters.
    /// </summary>
    internal string StringValue1
    {
        get
        {
            return this.GetXmlNodeString("@stringValue1");
        }
        set
        {
            this.SetXmlNodeString("@stringValue1", value, true);
        }
    }
    /// <summary>
    /// The string value 2 used by caption filters.
    /// </summary>
    internal string StringValue2
    {
        get
        {
            return this.GetXmlNodeString("@stringValue2");
        }
        set
        {
            this.SetXmlNodeString("@stringValue2", value, true);
        }
    }
    ExcelFilterColumn _filter = null;
    internal ExcelFilterColumn Filter
    {
        get
        {
            if (this._filter == null)
            {
                XmlNode? filterNode = this.GetNode("d:autoFilter/d:filterColumn");
                if (filterNode != null)
                {
                    switch (filterNode.LocalName)
                    {
                        case "customFilters":
                            this._filter = new ExcelCustomFilterColumn(this.NameSpaceManager, filterNode);
                            break;
                        case "top10":
                            this._filter = new ExcelTop10FilterColumn(this.NameSpaceManager, filterNode);
                            break;
                        case "filters":
                            this._filter = new ExcelValueFilterColumn(this.NameSpaceManager, filterNode);
                            break;
                        case "dynamicFilter":
                            this._filter = new ExcelDynamicFilterColumn(this.NameSpaceManager, filterNode);
                            break;
                        case "colorFilter":
                            this._filter = new ExcelColorFilterColumn(this.NameSpaceManager, filterNode);
                            break;
                        case "iconFilter":
                            this._filter = new ExcelIconFilterColumn(this.NameSpaceManager, filterNode);
                            break;
                        default:
                            this._filter = null;
                            break;
                    }
                }
                else
                {
                    throw new Exception("Invalid xml in pivot table. Missing Filter column");
                }
            }
            return this._filter;
        }
        set
        {
            this._filter = value;
        }
    }
}