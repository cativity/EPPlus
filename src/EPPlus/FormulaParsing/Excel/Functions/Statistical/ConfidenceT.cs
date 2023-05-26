/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "5.5",
        IntroducedInExcelVersion = "2010",
        Description = "Returns the confidence interval for a population mean, using a Student's t distribution")]
    internal class ConfidenceT : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            double alpha = this.ArgToDecimal(arguments, 0);
            double sigma = this.ArgToDecimal(arguments, 1);
            int size = this.ArgToInt(arguments, 2);

            if (alpha <= 0d || alpha >= 1d)
            {
                return this.CreateResult(eErrorType.Num);
            }

            if (sigma <= 0d)
            {
                return this.CreateResult(eErrorType.Num);
            }

            if (size < 1d)
            {
                return this.CreateResult(eErrorType.Num);
            }

            double result = System.Math.Abs(StudentInv(alpha / 2, size - 1) * sigma / System.Math.Sqrt(size));
            return this.CreateResult(result, DataType.Decimal);
        }


        private static double StudentInv(double p, double dof)
        {
            double x = BetaHelper.IBetaInv(2 * System.Math.Min(p, 1 - p), 0.5 * dof, 0.5);
            x = System.Math.Sqrt(dof * (1 - x) / x);
            return (p > 0.5) ? x : -x;
        }
    }
}
