﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "5.5",
        Description = "Converts a dollar price expressed as a fraction, into a dollar price expressed as a decimal")]
    internal class DollarDe : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            double fractionalDollar = ArgToDecimal(arguments, 0);
            double fractionDec = ArgToDecimal(arguments, 1);
            double fraction = System.Math.Floor(fractionDec);
            if (fraction <= 0d)
            {
                return this.CreateResult(eErrorType.Num);
            }

            if (fraction < 1d)
            {
                return this.CreateResult(eErrorType.Div0);
            }

            double intResult = System.Math.Floor(fractionalDollar);
            double result = ((double)intResult) + (fractionalDollar % 1) * System.Math.Pow(10d, (double)System.Math.Ceiling(System.Math.Log(fraction) / System.Math.Log(10))) / fraction;
            double power = System.Math.Pow(10d, (double)System.Math.Ceiling(System.Math.Log(fraction) / System.Math.Log(2)) + 1);
            return CreateResult(System.Math.Round(result * power) / power, DataType.Decimal);
        }
    }
}
