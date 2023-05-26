/*************************************************************************************************
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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Rounds a number towards zero, (i.e. rounds a positive number down and a negative number up), to a given number of digits")]
    internal class Rounddown : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            if (arguments.ElementAt(0).Value == null)
            {
                return this.CreateResult(0d, DataType.Decimal);
            }

            double number = this.ArgToDecimal(arguments, 0, context.Configuration.PrecisionAndRoundingStrategy);
            int nDecimals = this.ArgToInt(arguments, 1);

            int nFactor = number < 0 ? -1 : 1;
            number *= nFactor;

            double result;
            if (nDecimals > 0)
            {
                result = RoundDownDecimalNumber(number, nDecimals);
            }
            else
            {
                result = (int)System.Math.Floor(number);
                result -= (result % System.Math.Pow(10, (nDecimals*-1)));
            }
            return this.CreateResult(result * nFactor, DataType.Decimal);
        }

        private static double RoundDownDecimalNumber(double number, int nDecimals)
        {
            double integerPart = System.Math.Floor(number);
            double decimalPart = number - integerPart;
            decimalPart = System.Math.Pow(10d, nDecimals)*decimalPart;
            decimalPart = System.Math.Truncate(decimalPart)/System.Math.Pow(10d, nDecimals);
            double result = integerPart + decimalPart;
            return result;
        }
    }
}
