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
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

internal class ExcelLookupNavigator : LookupNavigator
{
    private int _currentRow;
    private int _currentCol;
    private object _currentValue;
    private RangeAddress _rangeAddress;
    private int _index;

    public ExcelLookupNavigator(LookupDirection direction, LookupArguments arguments, ParsingContext parsingContext)
        : base(direction, arguments, parsingContext) =>
        this.Initialize();

    private void Initialize()
    {
        this._index = 0;
        RangeAddressFactory? factory = new RangeAddressFactory(this.ParsingContext.ExcelDataProvider);

        if (this.Arguments.RangeInfo == null)
        {
            this._rangeAddress = factory.Create(this.ParsingContext.Scopes.Current.Address.Worksheet, this.Arguments.RangeAddress);
        }
        else
        {
            this._rangeAddress = factory.Create(this.Arguments.RangeInfo.Address.WorkSheetName, this.Arguments.RangeInfo.Address.Address);
        }

        this._currentCol = this._rangeAddress.FromCol;
        this._currentRow = this._rangeAddress.FromRow;
        this.SetCurrentValue();
    }

    private void SetCurrentValue() => this._currentValue = this.ParsingContext.ExcelDataProvider.GetCellValue(this._rangeAddress.Worksheet, this._currentRow, this._currentCol);

    private bool HasNext()
    {
        if (this.Direction == LookupDirection.Vertical)
        {
            return this._currentRow < this._rangeAddress.ToRow;
        }
        else
        {
            return this._currentCol < this._rangeAddress.ToCol;
        }
    }

    public override int Index => this._index;

    public override bool MoveNext()
    {
        if (!this.HasNext())
        {
            return false;
        }

        if (this.Direction == LookupDirection.Vertical)
        {
            this._currentRow++;
        }
        else
        {
            this._currentCol++;
        }

        this._index++;
        this.SetCurrentValue();

        return true;
    }

    public override object CurrentValue => this._currentValue;

    public override object GetLookupValue()
    {
        int row = this._currentRow;
        int col = this._currentCol;

        if (this.Direction == LookupDirection.Vertical)
        {
            col += this.Arguments.LookupIndex - 1;
            row += this.Arguments.LookupOffset;
        }
        else
        {
            row += this.Arguments.LookupIndex - 1;
            col += this.Arguments.LookupOffset;
        }

        return this.ParsingContext.ExcelDataProvider.GetCellValue(this._rangeAddress.Worksheet, row, col);
    }
}