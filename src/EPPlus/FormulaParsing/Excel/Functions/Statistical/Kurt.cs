﻿/*************************************************************************************************
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

using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;

[FunctionMetadata(Category = ExcelFunctionCategory.Statistical, EPPlusVersion = "6.0", Description = "Calculates the kurtosis of a data set")]
internal class Kurt : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        IEnumerable<ExcelDoubleCellValue>? numbers = this.ArgsToDoubleEnumerable(true, arguments, context, true);
        double n = (double)numbers.Count();

        if (n < 4)
        {
            return this.CreateResult(eErrorType.Div0);
        }

        double stdev = Stdev.StandardDeviation(numbers.Select(x => x.Value));

        if (stdev == 0d)
        {
            return this.CreateResult(eErrorType.Div0);
        }

        double part1 = n * (n + 1) / ((n - 1) * (n - 2) * (n - 3));
        double avg = numbers.Select(x => x.Value).Average();
        double part2 = 0d;

        for (int x = 0; x < n; x++)
        {
            part2 += System.Math.Pow(numbers.ElementAt(x) - avg, 4);
        }

        part2 /= System.Math.Pow(stdev, 4);
        double result = (part1 * part2) - (3 * System.Math.Pow(n - 1, 2) / ((n - 2) * (n - 3)));

        return this.CreateResult(result, DataType.Decimal);
    }
}