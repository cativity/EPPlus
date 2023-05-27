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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using System.Text.RegularExpressions;
using static OfficeOpenXml.FormulaParsing.EpplusExcelDataProvider;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

[FunctionMetadata(Category = ExcelFunctionCategory.LookupAndReference,
                  EPPlusVersion = "4",
                  Description =
                      "Searches for a specific value in one data vector, and returns a value from the corresponding position of a second data vector")]
internal class Lookup : LookupFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 2);

        if (HaveTwoRanges(arguments))
        {
            return this.HandleTwoRanges(arguments, context);
        }

        return this.HandleSingleRange(arguments, context);
    }

    private static bool HaveTwoRanges(IEnumerable<FunctionArgument> arguments)
    {
        if (arguments.Count() < 3)
        {
            return false;
        }

        return arguments.ElementAt(2).Value is RangeInfo;
    }

    private CompileResult HandleSingleRange(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        object? searchedValue = arguments.ElementAt(0).Value;
        Require.That(arguments.ElementAt(1).Value).Named("firstAddress").IsNotNull();
        string? firstAddress = ArgToAddress(arguments, 1, context);
        RangeAddressFactory? rangeAddressFactory = new RangeAddressFactory(context.ExcelDataProvider);
        RangeAddress? address = rangeAddressFactory.Create(firstAddress);
        int nRows = address.ToRow - address.FromRow;
        int nCols = address.ToCol - address.FromCol;
        int lookupIndex = nCols + 1;
        LookupDirection lookupDirection = LookupDirection.Vertical;

        if (nCols > nRows)
        {
            lookupIndex = nRows + 1;
            lookupDirection = LookupDirection.Horizontal;
        }

        LookupArguments? lookupArgs = new LookupArguments(searchedValue, firstAddress, lookupIndex, 0, true, arguments.ElementAt(1).ValueAsRangeInfo);
        LookupNavigator? navigator = LookupNavigatorFactory.Create(lookupDirection, lookupArgs, context);

        return this.Lookup(navigator, lookupArgs);
    }

    private CompileResult HandleTwoRanges(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        object? searchedValue = arguments.ElementAt(0).Value;
        Require.That(arguments.ElementAt(1).Value).Named("firstAddress").IsNotNull();
        Require.That(arguments.ElementAt(2).Value).Named("secondAddress").IsNotNull();
        string? firstAddress = ArgToAddress(arguments, 1, context);
        string? secondAddress = ArgToAddress(arguments, 2, context);
        RangeAddressFactory? rangeAddressFactory = new RangeAddressFactory(context.ExcelDataProvider);
        RangeAddress? address1 = rangeAddressFactory.Create(firstAddress);
        RangeAddress? address2 = rangeAddressFactory.Create(secondAddress);
        int lookupIndex = address2.FromCol - address1.FromCol + 1;
        int lookupOffset = address2.FromRow - address1.FromRow;
        LookupDirection lookupDirection = GetLookupDirection(address1);

        if (lookupDirection == LookupDirection.Horizontal)
        {
            lookupIndex = address2.FromRow - address1.FromRow + 1;
            lookupOffset = address2.FromCol - address1.FromCol;
        }

        LookupArguments? lookupArgs =
            new LookupArguments(searchedValue, firstAddress, lookupIndex, lookupOffset, true, arguments.ElementAt(1).ValueAsRangeInfo);

        LookupNavigator? navigator = LookupNavigatorFactory.Create(lookupDirection, lookupArgs, context);

        return this.Lookup(navigator, lookupArgs);
    }
}