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

using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;

[FunctionMetadata(Category = ExcelFunctionCategory.Statistical,
                  EPPlusVersion = "6.0",
                  Description = "Calculates the cumulative beta probability density function")]
internal class BetaDotDist : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 4);
        double x = this.ArgToDecimal(arguments, 0);
        double alpha = this.ArgToDecimal(arguments, 1);
        double beta = this.ArgToDecimal(arguments, 2);
        bool cumulative = this.ArgToBool(arguments, 3);
        double A = 0d;
        double B = 1d;

        if (arguments.Count() > 4)
        {
            A = this.ArgToDecimal(arguments, 4);
        }

        if (arguments.Count() > 5)
        {
            B = this.ArgToDecimal(arguments, 5);
        }

        // validate
        if (alpha <= 0 || beta <= 0)
        {
            return this.CreateResult(eErrorType.Num);
        }

        if (x < A || x > B || A == B)
        {
            return this.CreateResult(eErrorType.Num);
        }

        x = (x - A) / (B - A);
        double result = cumulative ? BetaHelper.BetaCdf(x, alpha, beta) : BetaHelper.BetaPdf(x, alpha, beta);

        return this.CreateResult(result, DataType.Decimal);
    }
}