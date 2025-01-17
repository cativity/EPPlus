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

using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sorting;

internal class TableSorter
{
    public TableSorter(ExcelTable table) => this._table = table;

    private readonly ExcelTable _table;

    public void Sort(TableSortOptions options) => this._table.DataRange.Sort(options, this._table);

    public void Sort(Action<TableSortOptions> configuration)
    {
        TableSortOptions? options = new TableSortOptions(this._table);
        configuration(options);
        this.Sort(options);
    }
}