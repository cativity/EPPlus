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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;

[FunctionMetadata(
                     Category = ExcelFunctionCategory.Financial,
                     EPPlusVersion = "6.0",
                     Description = "Calculates he accrued interest for a security that pays interest at maturity.")]
internal class AccrintM : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 4);
        // collect input
        System.DateTime issueDate = System.DateTime.FromOADate(this.ArgToInt(arguments, 0));
        System.DateTime settlementDate = System.DateTime.FromOADate(this.ArgToInt(arguments, 1));
        double rate = this.ArgToDecimal(arguments, 2);
        double par = this.ArgToDecimal(arguments, 3);
        int basis = 0;
        if (arguments.Count() > 4)
        {
            basis = this.ArgToInt(arguments, 4);
        }

        if (rate <= 0 || par <= 0)
        {
            return this.CreateResult(eErrorType.Num);
        }

        if (basis < 0 || basis > 4)
        {
            return this.CreateResult(eErrorType.Num);
        }

        if (issueDate >= settlementDate)
        {
            return this.CreateResult(eErrorType.Num);
        }

        DayCountBasis dayCountBasis = (DayCountBasis)basis;
        IFinanicalDays? fd = FinancialDaysFactory.Create(dayCountBasis);
        double result = fd.GetDaysBetweenDates(issueDate, settlementDate)/fd.DaysPerYear * rate * par;
        return this.CreateResult(result, DataType.Decimal);

    }
}