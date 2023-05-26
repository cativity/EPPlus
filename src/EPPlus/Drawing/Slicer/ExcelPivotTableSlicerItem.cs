using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Drawing.Slicer;

/// <summary>
/// Represents a pivot table slicer item.
/// </summary>
public class ExcelPivotTableSlicerItem
{
    private ExcelPivotTableSlicerCache _cache;
    private int _index;

    internal ExcelPivotTableSlicerItem(ExcelPivotTableSlicerCache cache, int index)
    {
        this._cache = cache;
        this._index = index;
    }
    /// <summary>
    /// The value of the item
    /// </summary>
    public object Value 
    { 
        get
        {
            if (this._index >= this._cache._field.Items.Count)
            {
                return null;
            }
            return this._cache._field.Items[this._index].Value;
        }
    }
    /// <summary>
    /// If the value is hidden 
    /// </summary>
    public bool Hidden 
    { 
        get
        {
            if (this._index >= this._cache._field.Items.Count)
            {
                throw(new IndexOutOfRangeException());
            }
            return this._cache._field.Items[this._index].Hidden;
        }
        set
        {
            if (this._index >= this._cache.Data.Items.Count)
            {
                throw (new IndexOutOfRangeException());
            }
            foreach (ExcelPivotTable? pt in this._cache.PivotTables)
            {
                ExcelPivotTableField? fld = pt.Fields[this._cache._field.Index];
                if (this._index >= fld.Items.Count || fld.Items[this._index].Type != eItemType.Data)
                {
                    continue;
                }

                fld.Items[this._index].Hidden = value;
            }
        }
    }
}