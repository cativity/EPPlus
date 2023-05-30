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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;

[FunctionMetadata(Category = ExcelFunctionCategory.Financial,
                  EPPlusVersion = "5.2",
                  Description = "Returns the sum-of-years' digits depreciation of an asset for a specified period")]
internal class Syd : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 4);
        double cost = this.ArgToDecimal(arguments, 0);
        double salvage = this.ArgToDecimal(arguments, 1);
        double life = this.ArgToDecimal(arguments, 2);
        double period = this.ArgToDecimal(arguments, 3);

        if (salvage < 0 || life <= 0 || period <= 0)
        {
            return this.CreateResult(eErrorType.Num);
        }

        double result = (cost - salvage) / (life * (life + 1));

        return this.CreateResult(result * (life + 1 - period) * 2, DataType.Decimal);
    }

    private static double GetInterest(double rate, double remainingAmount) => remainingAmount * rate;
}