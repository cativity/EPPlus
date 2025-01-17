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

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting;

/// <summary>
/// Sort options for sorting a range.
/// </summary>
public class RangeSortOptions : SortOptionsBase
{
    private RangeSortLayer _sortLayer;
    private RangeLeftToRightSortLayer _sortLayerLeftToRight;

    internal RangeSortOptions()
    {
    }

    /// <summary>
    /// Creates a new instance.
    /// </summary>
    /// <returns></returns>
    public static RangeSortOptions Create() => new();

    /// <summary>
    /// Creates the first sort layer (i.e. the first sort condition) for a row based/top to bottom sort.
    /// </summary>
    public RangeSortLayer SortBy => this._sortLayer ??= new RangeSortLayer(this);

    /// <summary>
    /// Creates the first sort layer (i.e. the first sort condition) for a column based/left to right sort.
    /// </summary>
    public RangeLeftToRightSortLayer SortLeftToRightBy => this._sortLayerLeftToRight ??= new RangeLeftToRightSortLayer(this);
}