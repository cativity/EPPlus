﻿/*************************************************************************************************
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

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

[FunctionMetadata(Category = ExcelFunctionCategory.Statistical,
                  EPPlusVersion = "5.5",
                  Description = "Returns the number of permutations for a given number of objects")]
internal class Permut : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 2);
        double number = this.ArgToDecimal(arguments, 0);
        number = System.Math.Floor(number);
        double numberChosen = this.ArgToDecimal(arguments, 1);

        if (number <= 0d || numberChosen <= 0 || number < numberChosen)
        {
            return this.CreateResult(eErrorType.Num);
        }

        double result = MathHelper.Factorial(number) / MathHelper.Factorial(number - numberChosen);

        return this.CreateResult(result, DataType.Decimal);
    }
}