﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/27/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;

internal abstract class CoupFunctionBase<T> : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 3);
        System.DateTime settlementDate = System.DateTime.FromOADate(this.ArgToInt(arguments, 0));
        System.DateTime maturityDate = System.DateTime.FromOADate(this.ArgToInt(arguments, 1));
        int frequency = this.ArgToInt(arguments, 2);
        int basis = 0;
        if (arguments.Count() >= 4)
        {
            basis = this.ArgToInt(arguments, 3);
        }
        // validate input
        if(settlementDate > maturityDate || (frequency != 1 && frequency != 2 && frequency != 4) || basis < 0 || basis > 4)
        {
            return this.CreateResult(eErrorType.Num);
        }
            
        FinanceCalcResult<T>? result = this.ExecuteFunction(FinancialDayFactory.Create(settlementDate, (DayCountBasis)basis), FinancialDayFactory.Create(maturityDate, (DayCountBasis)basis), frequency, (DayCountBasis)basis);
        if (result.HasError)
        {
            return this.CreateResult(result.ExcelErrorType);
        }

        if (typeof(T) == typeof(System.DateTime))
        {
            return this.CreateResult(Convert.ToDateTime(result.Result).ToOADate(), DataType.Date);
        }
        return this.CreateResult(result.Result, result.DataType);
    }

    protected abstract FinanceCalcResult<T> ExecuteFunction(FinancialDay settlementDate, FinancialDay maturityDate, int frequency, DayCountBasis basis = DayCountBasis.US_30_360);
}