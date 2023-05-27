/*************************************************************************************************
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

using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.LoadFunctions;

internal class LoadFromDictionaries : LoadFunctionBase
{
#if !NET35 && !NET40
    public LoadFromDictionaries(ExcelRangeBase range, IEnumerable<dynamic> items, LoadFromDictionariesParams parameters)
        : this(range, ConvertToDictionaries(items), parameters)
    {
    }
#endif

    public LoadFromDictionaries(ExcelRangeBase range, IEnumerable<IDictionary<string, object>> items, LoadFromDictionariesParams parameters)
        : base(range, parameters)
    {
        this._items = items;
        this._keys = parameters.Keys;
        this._headerParsingType = parameters.HeaderParsingType;
        this._cultureInfo = parameters.Culture ?? CultureInfo.CurrentCulture;

        if (items == null || !items.Any())
        {
            this._keys = Enumerable.Empty<string>();
        }
        else
        {
            IDictionary<string, object>? firstItem = items.First();

            if (this._keys == null || !this._keys.Any())
            {
                this._keys = firstItem.Keys;
            }
            else
            {
                this._keys = parameters.Keys;
            }
        }

        this._dataTypes = parameters.DataTypes ?? new eDataTypes[0];
    }

    private readonly IEnumerable<IDictionary<string, object>> _items;
    private readonly IEnumerable<string> _keys;
    private readonly eDataTypes[] _dataTypes;
    private readonly HeaderParsingTypes _headerParsingType;
    private readonly CultureInfo _cultureInfo;

#if !NET35 && !NET40
    private static IEnumerable<IDictionary<string, object>> ConvertToDictionaries(IEnumerable<dynamic> items)
    {
        List<Dictionary<string, object>>? result = new List<Dictionary<string, object>>();

        foreach (dynamic? item in items)
        {
            object? obj = item as object;

            if (obj != null)
            {
                Dictionary<string, object>? dict = new Dictionary<string, object>();
                MemberInfo[]? members = obj.GetType().GetMembers();

                foreach (MemberInfo? member in members)
                {
                    string? key = member.Name;
                    object value = null;

                    if (member is PropertyInfo)
                    {
                        value = ((PropertyInfo)member).GetValue(obj);
                        dict.Add(key, value);
                    }
                    else if (member is FieldInfo)
                    {
                        value = ((FieldInfo)member).GetValue(obj);
                        dict.Add(key, value);
                    }
                }

                if (dict.Count > 0)
                {
                    result.Add(dict);
                }
            }
        }

        return result;
    }

#endif

    protected override void LoadInternal(object[,] values, out Dictionary<int, FormulaCell> formulaCells, out Dictionary<int, string> columnFormats)
    {
        columnFormats = new Dictionary<int, string>();
        formulaCells = new Dictionary<int, FormulaCell>();

        int col = 0,
            row = 0;

        if (this.PrintHeaders && this._keys.Count() > 0)
        {
            foreach (string? key in this._keys)
            {
                values[row, col++] = this.ParseHeader(key);
            }

            row++;
        }

        foreach (IDictionary<string, object>? item in this._items)
        {
            col = 0;

            foreach (string? key in this._keys)
            {
                if (item.ContainsKey(key))
                {
                    object? v = item[key];

                    if (col < this._dataTypes.Length && v != null)
                    {
                        eDataTypes dataType = this._dataTypes[col];

                        switch (dataType)
                        {
                            case eDataTypes.Percent:
                            case eDataTypes.Number:
                                if (double.TryParse(v.ToString(), NumberStyles.Float | NumberStyles.Number, this._cultureInfo, out double d))
                                {
                                    if (dataType == eDataTypes.Percent)
                                    {
                                        d /= 100d;
                                    }

                                    values[row, col] = d;
                                }

                                break;

                            case eDataTypes.DateTime:
                                if (DateTime.TryParse(v.ToString(), out DateTime dt))
                                {
                                    values[row, col] = dt;
                                }

                                break;

                            case eDataTypes.String:
                                values[row, col] = v.ToString();

                                break;

                            default:
                                values[row, col] = v;

                                break;
                        }
                    }
                    else
                    {
                        values[row, col] = item[key];
                    }

                    col++;
                }
                else
                {
                    col++;
                }
            }

            row++;
        }
    }

    protected override int GetNumberOfRows()
    {
        if (this._items == null)
        {
            return 0;
        }

        return this._items.Count();
    }

    protected override int GetNumberOfColumns()
    {
        if (this._keys == null)
        {
            return 0;
        }

        return this._keys.Count();
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