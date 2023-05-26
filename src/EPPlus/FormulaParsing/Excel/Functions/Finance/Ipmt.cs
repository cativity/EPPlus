﻿/*************************************************************************************************
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

[FunctionMetadata(
                     Category = ExcelFunctionCategory.Financial,
                     EPPlusVersion = "5.2",
                     Description = "Calculates the interest payment for a given period of an investment, with periodic constant payments and a constant interest rate")]
internal class Ipmt : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 4);
        double rate = this.ArgToDecimal(arguments, 0);
        int per = this.ArgToInt(arguments, 1);
        int nPer = this.ArgToInt(arguments, 2);
        double presentValue = this.ArgToDecimal(arguments, 3);
        double fv = 0d;
        if (arguments.Count() >= 5)
        {
            fv = this.ArgToDecimal(arguments, 4);
        }
        PmtDue type = PmtDue.EndOfPeriod;
        if (arguments.Count() >= 6)
        {
            type = (PmtDue)this.ArgToInt(arguments, 5);
        }
        FinanceCalcResult<double>? result = IPmtImpl.Ipmt(rate, per, nPer, presentValue, fv, type);
        if (result.HasError)
        {
            return this.CreateResult(result.ExcelErrorType);
        }

        return this.CreateResult(result.Result, DataType.Decimal);
    }

    private static double GetInterest(double rate, double remainingAmount)
    {
        return remainingAmount * rate;
    }
}