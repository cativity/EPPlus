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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;

[FunctionMetadata(Category = ExcelFunctionCategory.Financial,
                  EPPlusVersion = "4",
                  Description = "Calculates the payments required to reduce a loan, from a supplied present value to a specified future value")]
internal class Pmt : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 3);
        double rate = this.ArgToDecimal(arguments, 0);
        int nPer = this.ArgToInt(arguments, 1);
        double presentValue = this.ArgToDecimal(arguments, 2);
        int payEndOfPeriod = 0;
        double futureValue = 0d;

        if (arguments.Count() > 3)
        {
            futureValue = this.ArgToDecimal(arguments, 3);
        }

        if (arguments.Count() > 4)
        {
            payEndOfPeriod = this.ArgToInt(arguments, 4);
        }

        FinanceCalcResult<double>? result =
            InternalMethods.PMT_Internal(rate, nPer, presentValue, futureValue, payEndOfPeriod == 0 ? PmtDue.EndOfPeriod : PmtDue.BeginningOfPeriod);

        if (result.HasError)
        {
            return this.CreateResult(result.ExcelErrorType);
        }

        return this.CreateResult(result.Result, DataType.Decimal);
    }

    private static double GetInterest(double rate, double remainingAmount) => remainingAmount * rate;
}