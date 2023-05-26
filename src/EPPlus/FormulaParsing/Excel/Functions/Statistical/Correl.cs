/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/10/2022         EPPlus Software AB           EPPlus v6
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using MathObj = System.Math;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;

[FunctionMetadata(
                     Category = ExcelFunctionCategory.Statistical,
                     EPPlusVersion = "6.0",
                     Description = "Returns the correlation coefficient of two cell ranges")]
internal class Correl : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 2);
        FunctionArgument? arg1 = arguments.ElementAt(0);
        FunctionArgument? arg2 = arguments.ElementAt(1);
        ExcelDoubleCellValue[]? arr1 = this.ArgsToDoubleEnumerable(new FunctionArgument[] { arg1 }, context).ToArray();
        ExcelDoubleCellValue[]? arr2 = this.ArgsToDoubleEnumerable(new FunctionArgument[] { arg2 }, context).ToArray();
        if (arr2.Count() != arr1.Count())
        {
            return this.CreateResult(eErrorType.NA);
        }

        if (arr1.Sum(x => x.Value) == 0 || arr2.Sum(x => x.Value) == 0)
        {
            return this.CreateResult(eErrorType.Div0);
        }

        double result = Covar(arr1, arr2) / StandardDeviation(arr1) / StandardDeviation(arr2);
        return this.CreateResult(result, DataType.Decimal);
    }

    private static double StandardDeviation(ExcelDoubleCellValue[] values)
    {
        double avg = values.Average(x => x.Value);
        return MathObj.Sqrt(values.Average(v => MathObj.Pow(v - avg, 2)));
    }

    private static double Covar(ExcelDoubleCellValue[] array1, ExcelDoubleCellValue[] array2)
    {
        double avg1 = array1.Select(x => x.Value).Average();
        double avg2 = array2.Select(x => x.Value).Average();
        double result = 0d;
        for (int x = 0; x < array1.Length; x++)
        {
            result += (array1[x] - avg1) * (array2[x] - avg2);
        }
        result /= array1.Length;
        return result;
    }
}