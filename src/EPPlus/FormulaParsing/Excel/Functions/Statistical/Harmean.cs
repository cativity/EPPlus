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

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;

[FunctionMetadata(Category = ExcelFunctionCategory.Statistical, EPPlusVersion = "6.0", Description = "Returns the harmonic mean of a data set.")]
internal class Harmean : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        IEnumerable<ExcelDoubleCellValue>? numbers = this.ArgsToDoubleEnumerable(arguments, context);

        if (numbers.Any(x => x.Value <= 0d))
        {
            return this.CreateResult(eErrorType.Num);
        }

        double n = 0d;

        for (int x = 0; x < numbers.Count(); x++)
        {
            n += 1d / numbers.ElementAt(x);
        }

        double result = (double)numbers.Count() / n;

        return this.CreateResult(result, DataType.Decimal);
    }
}