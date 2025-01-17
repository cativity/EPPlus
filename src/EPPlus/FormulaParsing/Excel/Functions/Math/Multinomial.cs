﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/10/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

[FunctionMetadata(Category = ExcelFunctionCategory.MathAndTrig,
                  EPPlusVersion = "5.5",
                  Description = "Returns the ratio of the factorial of a sum of values to the product of factorials.")]
internal class Multinomial : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        IEnumerable<ExcelDoubleCellValue>? numbers = this.ArgsToDoubleEnumerable(arguments, context);
        double part1 = 0d;
        double part2 = 1d;

        foreach (ExcelDoubleCellValue number in numbers)
        {
            part1 += number;
            part2 *= MathHelper.Factorial(number);
        }

        double result = MathHelper.Factorial(part1) / part2;

        return this.CreateResult(result, DataType.Decimal);
    }
}