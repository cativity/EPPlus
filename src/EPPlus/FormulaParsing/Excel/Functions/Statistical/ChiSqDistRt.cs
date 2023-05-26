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
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
            Category = ExcelFunctionCategory.Statistical,
            EPPlusVersion = "6.0",
            IntroducedInExcelVersion = "2010",
            Description = "Calculates the right-tailed probability of the Chi-Square Distribution")]
    internal class ChiSqDistRt : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            double n = this.ArgToDecimal(arguments, 0);
            int degreesOfFreedom = this.ArgToInt(arguments, 1);
            if(n < 0d || degreesOfFreedom < 1 || degreesOfFreedom > System.Math.Pow(10, 10))
            {
                return this.CreateResult(eErrorType.Num);
            }
            double result = 1d - ChiSquareHelper.CumulativeDistribution(n, degreesOfFreedom);
            return this.CreateResult(result, DataType.Decimal);
        }
    }
}
