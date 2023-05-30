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
                  EPPlusVersion = "5.2",
                  Description = "Calculates the cumulative principal paid on a loan, between two specified periods")]
internal class Cumprinc : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 6);
        double rate = this.ArgToDecimal(arguments, 0);
        double nPer = this.ArgToDecimal(arguments, 1);
        double presentValue = this.ArgToDecimal(arguments, 2);
        int startPeriod = this.ArgToInt(arguments, 3);
        int endPeriod = this.ArgToInt(arguments, 4);
        int t = this.ArgToInt(arguments, 5);

        if (t < 0 || t > 1)
        {
            return this.CreateResult(eErrorType.Num);
        }

        CumprincImpl? func = new CumprincImpl(new PmtProvider(), new FvProvider());
        FinanceCalcResult<double>? result = func.GetCumprinc(rate, nPer, presentValue, startPeriod, endPeriod, (PmtDue)t);

        if (result.HasError)
        {
            return this.CreateResult(result.ExcelErrorType);
        }

        return this.CreateResult(result.Result, DataType.Decimal);
    }

    private static double GetInterest(double rate, double remainingAmount) => remainingAmount * rate;
}