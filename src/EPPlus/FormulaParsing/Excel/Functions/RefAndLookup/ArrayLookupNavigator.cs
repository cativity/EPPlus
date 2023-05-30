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
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

internal class ArrayLookupNavigator : LookupNavigator
{
    private readonly FunctionArgument[] _arrayData;
    private int _index;
    //private object _currentValue;

    public ArrayLookupNavigator(LookupDirection direction, LookupArguments arguments, ParsingContext parsingContext)
        : base(direction, arguments, parsingContext)
    {
        Require.That(arguments).Named("arguments").IsNotNull();
        Require.That(arguments.DataArray).Named("arguments.DataArray").IsNotNull();
        this._arrayData = arguments.DataArray.ToArray();
        this.Initialize();
    }

    private void Initialize()
    {
        if (this.Arguments.LookupIndex >= this._arrayData.Length)
        {
            throw new ExcelErrorValueException(eErrorType.Ref);
        }

        //this.SetCurrentValue();
    }

    public override int Index
    {
        get { return this._index; }
    }

    //private void SetCurrentValue()
    //{
    //    this._currentValue = this._arrayData[this._index];
    //}

    private bool HasNext()
    {
        if (this.Direction == LookupDirection.Vertical)
        {
            return this._index < this._arrayData.Length - 1;
        }
        else
        {
            return false;
        }
    }

    public override bool MoveNext()
    {
        if (!this.HasNext())
        {
            return false;
        }

        if (this.Direction == LookupDirection.Vertical)
        {
            this._index++;
        }

        //this.SetCurrentValue();

        return true;
    }

    public override object CurrentValue
    {
        get { return this._arrayData[this._index].Value; }
    }

    public override object GetLookupValue()
    {
        return this._arrayData[this._index].Value;
    }
}