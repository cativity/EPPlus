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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;

[FunctionMetadata(
                     Category = ExcelFunctionCategory.Statistical,
                     EPPlusVersion = "5.5",
                     IntroducedInExcelVersion = "2010",
                     Description = "Returns the confidence interval for a population mean, using a normal distribution")]
internal class ConfidenceNorm : ExcelFunction
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

        double result = NormalCi(1, alpha, sigma ,size);
        result -= 1d;
        return this.CreateResult(result, DataType.Decimal);

    }

    private static double NormalCi(int s, double alpha, double sigma, int size)
    {
        double change = System.Math.Abs(NormalInv(alpha / 2, 0d, 1d) * sigma / System.Math.Sqrt((double)size));
        return 1d + change;
    }

    private static double NormalInv(double p, double mean, double std)
    {
        double n = (-1.41421356237309505 * std * ErfHelper.Erfcinv(2 * p)) + mean;
        return n;
    }

        
}