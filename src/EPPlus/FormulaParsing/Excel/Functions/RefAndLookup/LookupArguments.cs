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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

internal class LookupArguments
{
    public enum LookupArgumentDataType
    {
        ExcelRange,
        DataArray
    }

    public LookupArguments(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        : this(arguments, new ArgumentParsers(), context)
    {

    }

    public LookupArguments(IEnumerable<FunctionArgument> arguments, ArgumentParsers argumentParsers, ParsingContext context)
    {
        this._argumentParsers = argumentParsers;
        this.SearchedValue = arguments.ElementAt(0).Value;
        object? arg1 = arguments.ElementAt(1).Value;
        IEnumerable<FunctionArgument>? dataArray = arg1 as IEnumerable<FunctionArgument>;
        if (dataArray != null)
        {
            this.DataArray = dataArray;
            this.ArgumentDataType = LookupArgumentDataType.DataArray;
        }
        else
        {
            //if (arg1 is ExcelDataProvider.INameInfo) arg1 = ((ExcelDataProvider.INameInfo) arg1).Value;
            IRangeInfo? rangeInfo = arg1 as IRangeInfo;
            if (rangeInfo != null)
            {
                this.RangeAddress = string.IsNullOrEmpty(rangeInfo.Address.WorkSheetName) ? rangeInfo.Address.Address : "'" + rangeInfo.Address.WorkSheetName + "'!" + rangeInfo.Address.Address;
                this.RangeInfo = rangeInfo;
                this.ArgumentDataType = LookupArgumentDataType.ExcelRange;
            }
            else
            {
                this.RangeAddress = arg1.ToString();
                this.ArgumentDataType = LookupArgumentDataType.ExcelRange;
            }  
        }
        FunctionArgument? indexVal = arguments.ElementAt(2);

        if (indexVal.DataType == DataType.ExcelAddress)
        {
            ExcelAddress? address = new ExcelAddress(indexVal.Value.ToString());
            object? indexObj = context.ExcelDataProvider.GetRangeValue(address.WorkSheetName, address._fromRow, address._fromCol);
            this.LookupIndex = (int)this._argumentParsers.GetParser(DataType.Integer).Parse(indexObj);
        }
        else
        {
            this.LookupIndex = (int)this._argumentParsers.GetParser(DataType.Integer).Parse(arguments.ElementAt(2).Value);
        }
            
        if (arguments.Count() > 3)
        {
            this.RangeLookup = (bool)this._argumentParsers.GetParser(DataType.Boolean).Parse(arguments.ElementAt(3).Value);
        }
        else
        {
            this.RangeLookup = true;
        }
    }

    public LookupArguments(object searchedValue, string rangeAddress, int lookupIndex, int lookupOffset, bool rangeLookup, IRangeInfo rangeInfo)
    {
        this.SearchedValue = searchedValue;
        this.RangeAddress = rangeAddress;
        this.RangeInfo = rangeInfo;
        this.LookupIndex = lookupIndex;
        this.LookupOffset = lookupOffset;
        this.RangeLookup = rangeLookup;
    }

    private readonly ArgumentParsers _argumentParsers;

    public object SearchedValue { get; private set; }

    public string RangeAddress { get; private set; }

    public int LookupIndex { get; private set; }

    public int LookupOffset { get; private set; }

    public bool RangeLookup { get; private set; }

    public IEnumerable<FunctionArgument> DataArray { get; private set; }

    public IRangeInfo RangeInfo { get; private set; }

    public LookupArgumentDataType ArgumentDataType { get; private set; } 
}