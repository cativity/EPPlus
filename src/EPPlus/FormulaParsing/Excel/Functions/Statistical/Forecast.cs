/*************************************************************************************************
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

[FunctionMetadata(
                     Category = ExcelFunctionCategory.Statistical,
                     EPPlusVersion = "6.0",
                     Description = "Calculate, or predict, a future value by using existing values. The future value is a y-value for a given x-value.")]
internal class Forecast : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 3);
        double x = this.ArgToDecimal(arguments, 0);
        FunctionArgument? arg1 = arguments.ElementAt(1);
        FunctionArgument? arg2 = arguments.ElementAt(2);
        double[]? arrayY = this.ArgsToDoubleEnumerable(false, false, new FunctionArgument[] { arg1 }, context).Select(a => a.Value).ToArray();
        double[]? arrayX = this.ArgsToDoubleEnumerable(false, false, new FunctionArgument[] { arg2 }, context).Select(b => b.Value).ToArray();
        if (arrayY.Count() != arrayX.Count())
        {
            return this.CreateResult(eErrorType.NA);
        }

        if (!arrayY.Any())
        {
            return this.CreateResult(eErrorType.NA);
        }

        double result = ForecastImpl(x, arrayY, arrayX);
        return this.CreateResult(result, DataType.Decimal);
    }

    internal static double ForecastImpl(double x, double[] arrayY, double[] arrayX)
    {
        double avgY = arrayY.Average();
        double avgX = arrayX.Average();
        int nItems = arrayY.Length;
        double upperEquationPart = 0d;
        double lowerEquationPart = 0d;
        for (int ix = 0; ix < nItems; ix++)
        {
            upperEquationPart += (arrayX[ix] - avgX) * (arrayY[ix] - avgY);
            lowerEquationPart += System.Math.Pow(arrayX[ix] - avgX, 2);
        }
        double b = upperEquationPart / lowerEquationPart;
        double a = avgY - b * avgX;
        return a + b * x;
    }
}