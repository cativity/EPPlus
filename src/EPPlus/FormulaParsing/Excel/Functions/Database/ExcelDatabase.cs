/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

internal class ExcelDatabase
{
    private readonly ExcelDataProvider _dataProvider;
    private readonly int _fromCol;
    private readonly int _toCol;
    private readonly int _fieldRow;
    private readonly int _endRow;
    private readonly string _worksheet;
    private int _rowIndex;
    private readonly List<ExcelDatabaseField> _fields = new List<ExcelDatabaseField>();

    public IEnumerable<ExcelDatabaseField> Fields
    {
        get { return this._fields; }
    }

    public ExcelDatabase(ExcelDataProvider dataProvider, string range)
    {
        this._dataProvider = dataProvider;
        ExcelAddressBase? address = new ExcelAddressBase(range);
        this._fromCol = address._fromCol;
        this._toCol = address._toCol;
        this._fieldRow = address._fromRow;
        this._endRow = address._toRow;
        this._worksheet = address.WorkSheetName;
        this._rowIndex = this._fieldRow;
        this.Initialize();
    }

    private void Initialize()
    {
        int fieldIx = 0;

        for (int colIndex = this._fromCol; colIndex <= this._toCol; colIndex++)
        {
            object? nameObj = this.GetCellValue(this._fieldRow, colIndex);
            string? name = nameObj != null ? nameObj.ToString().ToLower(CultureInfo.InvariantCulture) : string.Empty;
            this._fields.Add(new ExcelDatabaseField(name, fieldIx++));
        }
    }

    private object GetCellValue(int row, int col)
    {
        return this._dataProvider.GetRangeValue(this._worksheet, row, col);
    }

    public bool HasMoreRows
    {
        get { return this._rowIndex < this._endRow; }
    }

    public ExcelDatabaseRow Read()
    {
        ExcelDatabaseRow? retVal = new ExcelDatabaseRow();
        this._rowIndex++;

        foreach (ExcelDatabaseField? field in this.Fields)
        {
            int colIndex = this._fromCol + field.ColIndex;
            object? val = this.GetCellValue(this._rowIndex, colIndex);
            retVal[field.FieldName] = val;
        }

        return retVal;
    }
}