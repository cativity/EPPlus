﻿/*************************************************************************************************
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
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable.Filter;

/// <summary>
/// The base collection for pivot table filters
/// </summary>
public abstract class ExcelPivotTableFilterBaseCollection : IEnumerable<ExcelPivotTableFilter>
{
    internal List<ExcelPivotTableFilter> _filters = new List<ExcelPivotTableFilter>();
    internal readonly ExcelPivotTable _table;
    internal readonly ExcelPivotTableField _field;

    internal ExcelPivotTableFilterBaseCollection(ExcelPivotTable table)
    {
        this._table = table;
        XmlNode? filtersNode = this._table.GetNode("d:filters");

        if (filtersNode != null)
        {
            foreach (XmlNode node in filtersNode.ChildNodes)
            {
                ExcelPivotTableFilter? f = new ExcelPivotTableFilter(this._table.NameSpaceManager, node, this._table.WorkSheet.Workbook.Date1904);
                table.SetNewFilterId(f.Id);
                this._filters.Add(f);
            }
        }
    }

    internal ExcelPivotTableFilterBaseCollection(ExcelPivotTableField field)
    {
        this._field = field;
        this._table = field._pivotTable;

        foreach (ExcelPivotTableFilter? filter in this._table.Filters)
        {
            if (filter.Fld == field.Index)
            {
                this._filters.Add(filter);
            }
        }
    }

    /// <summary>
    /// Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    public IEnumerator<ExcelPivotTableFilter> GetEnumerator() => this._filters.GetEnumerator();

    /// <summary>
    /// Returns an enumerator that iterates through the collection.
    /// </summary>
    /// <returns>An enumerator that can be used to iterate through the collection.</returns>
    IEnumerator IEnumerable.GetEnumerator() => this._filters.GetEnumerator();

    internal XmlNode GetOrCreateFiltersNode() => this._table.CreateNode("d:filters");

    internal ExcelPivotTableFilter CreateFilter()
    {
        XmlNode? topNode = this.GetOrCreateFiltersNode();
        XmlElement? filterNode = topNode.OwnerDocument.CreateElement("filter", ExcelPackage.schemaMain);
        _ = topNode.AppendChild(filterNode);

        ExcelPivotTableFilter? filter = new ExcelPivotTableFilter(this._field.NameSpaceManager, filterNode, this._table.WorkSheet.Workbook.Date1904)
        {
            EvalOrder = -1, Fld = this._field.Index, Id = this._table.GetNewFilterId()
        };

        return filter;
    }

    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count => this._filters.Count;

    /// <summary>
    /// The indexer for the collection
    /// </summary>
    /// <param name="index">The index</param>
    /// <returns></returns>
    public ExcelPivotTableFilter this[int index]
    {
        get
        {
            if (index < 0 || index >= this._filters.Count)
            {
                throw new ArgumentOutOfRangeException();
            }

            return this._filters[index];
        }
    }
}