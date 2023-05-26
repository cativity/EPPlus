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
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

internal class ExcelDatabaseCriteria
{
    private readonly ExcelDataProvider _dataProvider;
    private readonly int _fromCol;
    private readonly int _toCol;
    private readonly string _worksheet;
    private readonly int _fieldRow;
    private readonly Dictionary<ExcelDatabaseCriteriaField, object> _criterias = new Dictionary<ExcelDatabaseCriteriaField, object>();

    internal ExcelDatabaseCriteria(ExcelDataProvider dataProvider, string range)
    {
        this._dataProvider = dataProvider;
        ExcelAddressBase? address = new ExcelAddressBase(range);
        this._fromCol = address._fromCol;
        this._toCol = address._toCol;
        this._worksheet = address.WorkSheetName;
        this._fieldRow = address._fromRow;
        this.Initialize();
    }

    private void Initialize()
    {
        for (int x = this._fromCol; x <= this._toCol; x++)
        {
            object? fieldObj = this._dataProvider.GetCellValue(this._worksheet, this._fieldRow, x);
            object? val = this._dataProvider.GetCellValue(this._worksheet, this._fieldRow + 1, x);
            if (fieldObj != null && val != null)
            {
                if(fieldObj is string)
                { 
                    ExcelDatabaseCriteriaField? field = new ExcelDatabaseCriteriaField(fieldObj.ToString().ToLower(CultureInfo.InvariantCulture));
                    this._criterias.Add(field, val);
                }
                else if (ConvertUtil.IsNumericOrDate(fieldObj))
                {
                    ExcelDatabaseCriteriaField? field = new ExcelDatabaseCriteriaField((int) fieldObj);
                    this._criterias.Add(field, val);
                }

            }
        }
    }

    public virtual IDictionary<ExcelDatabaseCriteriaField, object> Items
    {
        get { return this._criterias; }
    }
}