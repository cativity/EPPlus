﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/16/2020         EPPlus Software AB       EPPlus 5.2.1
 *************************************************************************************************/

using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using OfficeOpenXml.LoadFunctions.Params;
using System.Linq.Expressions;
using System.Text.RegularExpressions;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.LoadFunctions;

internal class LoadFromCollection<T> : LoadFunctionBase
{
    public LoadFromCollection(ExcelRangeBase range, IEnumerable<T> items, LoadFromCollectionParams parameters)
        : base(range, parameters)
    {
        this._items = items;
        this._bindingFlags = parameters.BindingFlags;
        this._headerParsingType = parameters.HeaderParsingType;
        Type? type = typeof(T);
        EpplusTableAttribute? tableAttr = type.GetFirstAttributeOfType<EpplusTableAttribute>();

        if (tableAttr != null)
        {
            this.ShowFirstColumn = tableAttr.ShowFirstColumn;
            this.ShowLastColumn = tableAttr.ShowLastColumn;
            this.ShowTotal = tableAttr.ShowTotal;
        }

        EPPlusTableColumnSortOrderAttribute? classSortOrderAttr = type.GetFirstAttributeOfType<EPPlusTableColumnSortOrderAttribute>();

        if (classSortOrderAttr != null && classSortOrderAttr.Properties != null && classSortOrderAttr.Properties.Length > 0)
        {
            this.SortOrderProperties = classSortOrderAttr.Properties.ToList();
        }

        if (parameters.Members == null)
        {
            LoadFromCollectionColumns<T>? cols = new LoadFromCollectionColumns<T>(parameters.BindingFlags, this.SortOrderProperties);
            List<ColumnInfo>? columns = cols.Setup();
            this._columns = columns.ToArray();
        }
        else
        {
            this._columns = parameters.Members.Select(x => new ColumnInfo { MemberInfo = x }).ToArray();

            if (this._columns.Length == 0) //Fixes issue 15555
            {
                throw new ArgumentException("Parameter Members must have at least one property. Length is zero");
            }

            foreach (ColumnInfo? columnInfo in this._columns)
            {
                if (columnInfo.MemberInfo == null)
                {
                    continue;
                }

                MemberInfo? member = columnInfo.MemberInfo;

                if (member.DeclaringType != null && member.DeclaringType != type)
                {
                    this._isSameType = false;
                }

                //Fixing inverted check for IsSubclassOf / Pullrequest from tom dam
                if (member.DeclaringType != null
                    && member.DeclaringType != type
                    && !TypeCompat.IsSubclassOf(type, member.DeclaringType)
                    && !TypeCompat.IsSubclassOf(member.DeclaringType, type))
                {
                    throw new InvalidCastException("Supplied properties in parameter Properties must be of the same type as T (or an assignable type from T)");
                }
            }
        }
    }

    private readonly BindingFlags _bindingFlags;
    private readonly ColumnInfo[] _columns;
    private readonly HeaderParsingTypes _headerParsingType;
    private readonly IEnumerable<T> _items;
    private readonly bool _isSameType = true;

    internal List<string> SortOrderProperties { get; private set; }

    protected override int GetNumberOfColumns() => this._columns.Length == 0 ? 1 : this._columns.Length;

    protected override int GetNumberOfRows()
    {
        if (this._items == null)
        {
            return 0;
        }

        return this._items.Count();
    }

    protected override void PostProcessTable(ExcelTable table, ExcelRangeBase range)
    {
        for (int ix = 0; ix < table.Columns.Count; ix++)
        {
            if (ix >= this._columns.Length)
            {
                break;
            }

            string? totalsRowFormula = this._columns[ix].TotalsRowFormula;
            string? totalsRowLabel = this._columns[ix].TotalsRowLabel;

            if (!string.IsNullOrEmpty(totalsRowFormula))
            {
                table.Columns[ix].TotalsRowFormula = totalsRowFormula;
            }
            else if (!string.IsNullOrEmpty(totalsRowLabel))
            {
                table.Columns[ix].TotalsRowLabel = this._columns[ix].TotalsRowLabel;
                table.Columns[ix].TotalsRowFunction = RowFunctions.None;
            }
            else
            {
                table.Columns[ix].TotalsRowFunction = this._columns[ix].TotalsRowFunction;
            }

            if (!string.IsNullOrEmpty(this._columns[ix].TotalsRowNumberFormat))
            {
                int row = range._toRow + 1;
                int col = range._fromCol + this._columns[ix].Index;
                range.Worksheet.Cells[row, col].Style.Numberformat.Format = this._columns[ix].TotalsRowNumberFormat;
            }
        }
    }

    protected override void LoadInternal(object[,] values, out Dictionary<int, FormulaCell> formulaCells, out Dictionary<int, string> columnFormats)
    {
        int col = 0,
            row = 0;

        columnFormats = new Dictionary<int, string>();
        formulaCells = new Dictionary<int, FormulaCell>();

        if (this._columns.Length > 0 && this.PrintHeaders)
        {
            this.SetHeaders(values, columnFormats, ref col, ref row);
        }

        if (!this._items.Any() && (this._columns.Length == 0 || this.PrintHeaders == false))
        {
            return;
        }

        this.SetValuesAndFormulas(values, formulaCells, ref col, ref row);
    }

    private void SetValuesAndFormulas(object[,] values, Dictionary<int, FormulaCell> formulaCells, ref int col, ref int row)
    {
        this.GetNumberOfColumns();

        foreach (T? item in this._items)
        {
            if (item == null)
            {
                col = this.GetNumberOfColumns();
            }
            else
            {
                col = 0;
                Type? t = item.GetType();

                if (item is string || item is decimal || item is DateTime || t.IsPrimitive)
                {
                    values[row, col++] = item;
                }
                else if (t.IsEnum)
                {
                    values[row, col++] = GetEnumValue(item, t);
                    ;
                }
                else
                {
                    foreach (ColumnInfo? colInfo in this._columns)
                    {
                        if (!string.IsNullOrEmpty(colInfo.Path) && colInfo.Path.Contains("."))
                        {
                            values[row, col++] = GetValueByPath(item, colInfo.Path);

                            continue;
                        }

                        T? obj = item;

                        if (colInfo.MemberInfo != null)
                        {
                            MemberInfo? member = colInfo.MemberInfo;
                            object v = null;

                            if (this._isSameType == false && obj.GetType().GetMember(member.Name, this._bindingFlags).Length == 0)
                            {
                                col++;

                                continue; //Check if the property exists if and inherited class is used
                            }
                            else if (member is PropertyInfo)
                            {
                                v = ((PropertyInfo)member).GetValue(obj, null);
                            }
                            else if (member is FieldInfo)
                            {
                                v = ((FieldInfo)member).GetValue(obj);
                            }
                            else if (member is MethodInfo)
                            {
                                v = ((MethodInfo)member).Invoke(obj, null);
                            }

#if (!NET35)
                            if (v != null)
                            {
                                Type? type = v.GetType();

                                if (type.IsEnum)
                                {
                                    v = GetEnumValue(v, type);
                                }
                            }
#endif

                            values[row, col++] = v;
                        }
                        else if (!string.IsNullOrEmpty(colInfo.Formula))
                        {
                            formulaCells[colInfo.Index] = new FormulaCell { Formula = colInfo.Formula };
                        }
                        else if (!string.IsNullOrEmpty(colInfo.FormulaR1C1))
                        {
                            formulaCells[colInfo.Index] = new FormulaCell { FormulaR1C1 = colInfo.FormulaR1C1 };
                        }
                    }
                }
            }

            row++;
        }
    }

    private static string GetEnumValue(object item, Type t)
    {
#if (NET35)
            return item.ToString();
#else
        string? v = item.ToString();
        MemberInfo? m = t.GetMember(v).FirstOrDefault();
        DescriptionAttribute? da = m.GetCustomAttribute<DescriptionAttribute>();

        return da?.Description ?? v;
#endif
    }

    private static object GetValueByPath(object obj, string path)
    {
        string[]? members = path.Split('.');
        object o = obj;

        foreach (string? member in members)
        {
            if (o == null)
            {
                return null;
            }

            MemberInfo[]? memberInfos = o.GetType().GetMember(member);

            if (memberInfos == null || memberInfos.Length == 0)
            {
                return null;
            }

            MemberInfo? memberInfo = memberInfos.First();

            if (memberInfo is PropertyInfo pi)
            {
                o = pi.GetValue(o, null);
            }
            else if (memberInfo is FieldInfo fi)
            {
                o = fi.GetValue(obj);
            }
            else if (memberInfo is MethodInfo mi)
            {
                o = mi.Invoke(obj, null);
            }
            else
            {
                throw new NotSupportedException("Invalid member: '" + memberInfo.Name + "', not supported member type '" + memberInfo.GetType().FullName + "'");
            }
        }

        return o;
    }

    private void SetHeaders(object[,] values, Dictionary<int, string> columnFormats, ref int col, ref int row)
    {
        foreach (ColumnInfo? colInfo in this._columns)
        {
            string? header = colInfo.Header;

            // if the header is already set and contains a space it doesn't need more formatting or validation.
            bool useExistingHeader = !string.IsNullOrEmpty(header) && header.Contains(" ");

            if (colInfo.MemberInfo != null)
            {
                // column data based on a property read with reflection
                MemberInfo? member = colInfo.MemberInfo;
                EpplusTableColumnAttribute? epplusColumnAttribute = member.GetFirstAttributeOfType<EpplusTableColumnAttribute>();

                if (epplusColumnAttribute != null)
                {
                    if (!useExistingHeader)
                    {
                        if (!string.IsNullOrEmpty(epplusColumnAttribute.Header))
                        {
                            header = epplusColumnAttribute.Header;
                        }
                        else
                        {
                            header = this.ParseHeader(member.Name);
                        }
                    }

                    if (!string.IsNullOrEmpty(epplusColumnAttribute.NumberFormat))
                    {
                        columnFormats.Add(col, epplusColumnAttribute.NumberFormat);
                    }
                }
                else if (!useExistingHeader)
                {
                    DescriptionAttribute? descriptionAttribute = member.GetFirstAttributeOfType<DescriptionAttribute>();

                    if (descriptionAttribute != null)
                    {
                        header = descriptionAttribute.Description;
                    }
                    else
                    {
                        DisplayNameAttribute? displayNameAttribute = member.GetFirstAttributeOfType<DisplayNameAttribute>();

                        if (displayNameAttribute != null)
                        {
                            header = displayNameAttribute.DisplayName;
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(colInfo.Header) && colInfo.Header != member.Name)
                            {
                                header = colInfo.Header;
                            }
                            else
                            {
                                header = this.ParseHeader(member.Name);
                            }
                        }
                    }
                }
            }
            else
            {
                // column is a FormulaColumn
                header = colInfo.Header;
                columnFormats.Add(colInfo.Index, colInfo.NumberFormat);
            }

            values[row, col++] = header;
        }

        row++;
    }

    private string ParseHeader(string header)
    {
        switch (this._headerParsingType)
        {
            case HeaderParsingTypes.Preserve:
                return header;

            case HeaderParsingTypes.UnderscoreToSpace:
                return header.Replace("_", " ");

            case HeaderParsingTypes.CamelCaseToSpace:
                return Regex.Replace(header, "([A-Z])", " $1", RegexOptions.Compiled).Trim();

            case HeaderParsingTypes.UnderscoreAndCamelCaseToSpace:
                header = Regex.Replace(header, "([A-Z])", " $1", RegexOptions.Compiled).Trim();

                return header.Replace("_ ", "_").Replace("_", " ");

            default:
                return header;
        }
    }
}