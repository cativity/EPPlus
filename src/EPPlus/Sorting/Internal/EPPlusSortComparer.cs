﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/07/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/

using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting.Internal;

internal class EPPlusSortComparer : EPPlusSortComparerBase<SortItem<ExcelValue>, ExcelValue>
{
    public EPPlusSortComparer(int[] columns,
                              bool[] descending,
                              Dictionary<int, string[]> customLists,
                              CultureInfo culture = null,
                              CompareOptions compareOptions = CompareOptions.None)
        : base(descending, customLists, culture, compareOptions) =>
        this._columns = columns;

    private readonly int[] _columns;

    public override int Compare(SortItem<ExcelValue> x, SortItem<ExcelValue> y)
    {
        for (int i = 0; i < this._columns.Length; i++)
        {
            object? x1 = x.Items[this._columns[i]]._value;
            object? y1 = y.Items[this._columns[i]]._value;

            if (x1 == null && y1 != null)
            {
                return 1;
            }

            if (x1 != null && y1 == null)
            {
                return -1;
            }

            int ret;

            if (this.CustomLists != null && this.CustomLists.ContainsKey(this._columns[i]))
            {
                int weight1 = this.GetSortWeightByCustomList(x1.ToString(), this.CustomLists[this._columns[i]]);
                int weight2 = this.GetSortWeightByCustomList(y1.ToString(), this.CustomLists[this._columns[i]]);

                if (weight1 != CustomListNotFound && weight2 != CustomListNotFound)
                {
                    ret = weight1.CompareTo(weight2);
                }
                else if (weight1 == CustomListNotFound && weight1 != weight2)
                {
                    return 1;
                }
                else if (weight2 == CustomListNotFound && weight1 != weight2)
                {
                    return -1;
                }
                else
                {
                    ret = CompareObjects(x1, y1);
                }
            }
            else
            {
                ret = CompareObjects(x1, y1);
            }

            if (ret != 0)
            {
                return ret * (this.Descending[i] ? -1 : 1);
            }
        }

        return 0;
    }
}