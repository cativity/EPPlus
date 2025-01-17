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

/// <summary>
/// Sort options for sorting an <see cref="ExcelTable"/>
/// </summary>
public class TableSortOptions : SortOptionsBase
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="table">The table sort</param>
    public TableSortOptions(ExcelTable table)
        : base()
    {
        this._table = table;
        this._columnNameIndexes = new Dictionary<string, int>();

        for (int x = 0; x < table.Columns.Count(); x++)
        {
            this._columnNameIndexes[table.Columns.ElementAt(x).Name] = x;
        }
    }

    private TableSortLayer _sortLayer;
    private readonly ExcelTable _table;
    private readonly Dictionary<string, int> _columnNameIndexes;

    internal ExcelTable Table => this._table;

    internal int GetColumnNameIndex(string name)
    {
        if (!this._columnNameIndexes.ContainsKey(name))
        {
            throw new InvalidOperationException($"Table {this._table.Name} does not contain column {name}");
        }

        return this._columnNameIndexes[name];
    }

    /// <summary>
    /// Defines the first <see cref="TableSortLayer"/>.
    /// </summary>
    public TableSortLayer SortBy => this._sortLayer ??= new TableSortLayer(this);
}