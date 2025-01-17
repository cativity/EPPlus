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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

[FunctionMetadata(Category = ExcelFunctionCategory.MathAndTrig,
                  EPPlusVersion = "4",
                  Description = "Rounds a number away from zero (i.e. rounds a positive number up and a negative number down), to a given number of digits")]
internal class Roundup : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 2);

        if (arguments.ElementAt(0).Value == null)
        {
            return this.CreateResult(0d, DataType.Decimal);
        }

        double number = this.ArgToDecimal(arguments, 0, context.Configuration.PrecisionAndRoundingStrategy);
        int nDigits = this.ArgToInt(arguments, 1);

        double result = number >= 0
                            ? System.Math.Ceiling(number * System.Math.Pow(10, nDigits)) / System.Math.Pow(10, nDigits)
                            : System.Math.Floor(number * System.Math.Pow(10, nDigits)) / System.Math.Pow(10, nDigits);

        return this.CreateResult(result, DataType.Decimal);
    }
}