﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

[FunctionMetadata(
                     Category = ExcelFunctionCategory.Statistical,
                     EPPlusVersion = "5.2",
                     Description = "Returns the K'th percentile of values in a supplied range, where K is in the range 0 - 1 (inclusive)")]
internal class Percentile : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 2);
        List<double>? arr = this.ArgsToDoubleEnumerable(arguments.Take(1), context).Select(x => (double)x).ToList();
        double percentile = this.ArgToDecimal(arguments, 1);
        if (percentile < 0 || percentile > 1)
        {
            return this.CreateResult(eErrorType.Num);
        }

        arr.Sort();
        int nElements = arr.Count;
        double dIx = percentile * (nElements - 1);
        int ix = (int)dIx;
        double rest = dIx - ix;
        double result = ix < (nElements - 1) ? arr[ix] + (arr[ix + 1] - arr[ix]) * rest : arr.Last();
        return this.CreateResult(result, DataType.Decimal);
    }
}