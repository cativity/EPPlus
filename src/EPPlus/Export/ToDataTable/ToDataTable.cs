/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/15/2020         EPPlus Software AB       ToDataTable function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable;

internal class ToDataTable
{
    public ToDataTable(ToDataTableOptions options, ExcelRangeBase range)
    {
        Require.That(options).IsNotNull();
        Require.That(range).IsNotNull();
        this._options = options;
        this._range = range;
    }

    private readonly ToDataTableOptions _options;
    private readonly ExcelRangeBase _range;

    public DataTable Execute()
    {
        DataTable? dataTable = new DataTableBuilder(this._options, this._range).Build();
        new DataTableExporter(this._options, this._range, dataTable).Export();
        return dataTable;
    }

    public DataTable Execute(DataTable dataTable)
    {
        Require.That(dataTable).IsNotNull();
        new DataTableMapper(this._options, this._range, dataTable).Map();
        new DataTableExporter(this._options, this._range, dataTable).Export();
        return dataTable;
    }
}