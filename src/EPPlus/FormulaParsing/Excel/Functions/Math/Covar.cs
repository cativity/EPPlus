﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/12/2020         EPPlus Software AB       Version 5.5
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

[FunctionMetadata(Category = ExcelFunctionCategory.Statistical,
                  EPPlusVersion = "5.5",
                  Description = "Returns covariance, the average of the products of deviations for each data point pair in two data sets.")]
internal class Covar : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 2);
        ExcelDoubleCellValue[]? array1 = this.ArgsToDoubleEnumerable(arguments.Take(1), context).ToArray();
        ExcelDoubleCellValue[]? array2 = this.ArgsToDoubleEnumerable(arguments.Skip(1).Take(1), context).ToArray();

        if (array1.Length != array2.Length)
        {
            return this.CreateResult(eErrorType.NA);
        }

        if (array1.Length == 0)
        {
            return this.CreateResult(eErrorType.Div0);
        }

        double avg1 = array1.Select(x => x.Value).Average();
        double avg2 = array2.Select(x => x.Value).Average();
        double result = 0d;

        for (int x = 0; x < array1.Length; x++)
        {
            result += (array1[x] - avg1) * (array2[x] - avg2);
        }

        result /= array1.Length;

        return this.CreateResult(result, DataType.Decimal);
    }
}