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
using System.Data;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.LoadFunctions;

internal class LoadFromDataTable
{
    public LoadFromDataTable(ExcelRangeBase range, DataTable dataTable, LoadFromDataTableParams parameters)
    {
        this._range = range;
        this._worksheet = range.Worksheet;
        this._dataTable = dataTable;
        this._printHeaders = parameters.PrintHeaders;
        this._tableStyle = parameters.TableStyle;
    }

    private readonly ExcelRangeBase _range;
    private readonly ExcelWorksheet _worksheet;
    private readonly DataTable _dataTable;
    private readonly bool _printHeaders;
    private TableStyles? _tableStyle;

    public ExcelRangeBase Load()
    {
        if (this._dataTable == null)
        {
            throw new ArgumentNullException("Table can't be null");
        }

        if (this._dataTable.Rows.Count == 0 && this._printHeaders == false)
        {
            return null;
        }

        //var rowArray = new List<object[]>();
        int row = this._range._fromRow;

        if (this._printHeaders)
        {
            this._worksheet._values.SetValueRow_Value(this._range._fromRow,
                                                      this._range._fromCol,
                                                      this._dataTable.Columns.Cast<DataColumn>().Select((dc) => { return dc.Caption; }).ToArray());

            row++;
        }

        foreach (DataRow dr in this._dataTable.Rows)
        {
            this._range.Worksheet._values.SetValueRow_Value(row++, this._range._fromCol, dr.ItemArray);
        }

        if (row != this._range._fromRow)
        {
            row--;
        }

        // set table style
        int rows = (this._dataTable.Rows.Count == 0 ? 1 : this._dataTable.Rows.Count) + (this._printHeaders ? 1 : 0);

        if (rows >= 0 && this._dataTable.Columns.Count > 0 && this._tableStyle.HasValue)
        {
            ExcelTable? tbl = this._worksheet.Tables.Add(new ExcelAddressBase(this._range._fromRow,
                                                                              this._range._fromCol,
                                                                              this._range._fromRow + rows - 1,
                                                                              this._range._fromCol + this._dataTable.Columns.Count - 1),
                                                         this._dataTable.TableName);

            tbl.ShowHeader = this._printHeaders;
            tbl.TableStyle = this._tableStyle.Value;
        }

        return this._worksheet.Cells[this._range._fromRow, this._range._fromCol, row, this._range._fromCol + this._dataTable.Columns.Count - 1];
    }
}