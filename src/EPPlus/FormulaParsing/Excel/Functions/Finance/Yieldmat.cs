/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/10/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;

[FunctionMetadata(
                     Category = ExcelFunctionCategory.Financial,
                     EPPlusVersion = "5.5",
                     Description = "Returns the annual yield of a security that pays interest at maturity.")]
internal class Yieldmat : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 5);
        System.DateTime settlementDate = System.DateTime.FromOADate(this.ArgToInt(arguments, 0));
        System.DateTime maturityDate = System.DateTime.FromOADate(this.ArgToInt(arguments, 1));
        if (settlementDate >= maturityDate)
        {
            return this.CreateResult(eErrorType.Num);
        }

        System.DateTime issueDate = System.DateTime.FromOADate(this.ArgToInt(arguments, 2));
            
        double rate = this.ArgToDecimal(arguments, 3);
        if (rate < 0)
        {
            return this.CreateResult(eErrorType.Num);
        }

        double price = this.ArgToDecimal(arguments, 4);
        if (price <= 0)
        {
            return this.CreateResult(eErrorType.Num);
        }

        int basis = 0;
        if(arguments.Count() > 5)
        {
            basis = this.ArgToInt(arguments, 5);
            if (basis < 0 || basis > 4)
            {
                return this.CreateResult(eErrorType.Num);
            }
        }

        YearFracProvider? yearFracProvider = new YearFracProvider(context);
        double yf1 = yearFracProvider.GetYearFrac(issueDate, maturityDate, (DayCountBasis)basis);
        double yf2 = yearFracProvider.GetYearFrac(issueDate, settlementDate, (DayCountBasis)basis);
        double yf3 = yearFracProvider.GetYearFrac(settlementDate, maturityDate, (DayCountBasis)basis);

        double result = 1d + yf1 * rate;
        result /= price / 100d + yf2 * rate;
        result = --result / yf3;
        return this.CreateResult(result, DataType.Decimal);
    }
}