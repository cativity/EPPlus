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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;

[FunctionMetadata(Category = ExcelFunctionCategory.Financial,
                  EPPlusVersion = "5.2",
                  Description = "Calculates the price per $100 face value of a security that pays periodic interest")]
internal class Price : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 6);
        System.DateTime settlementDate = System.DateTime.FromOADate(this.ArgToInt(arguments, 0));
        System.DateTime maturityDate = System.DateTime.FromOADate(this.ArgToInt(arguments, 1));
        double rate = this.ArgToDecimal(arguments, 2);
        double yield = this.ArgToDecimal(arguments, 3);
        double redemption = this.ArgToDecimal(arguments, 4);
        int frequency = this.ArgToInt(arguments, 5);
        int basis = 0;

        if (arguments.Count() >= 7)
        {
            basis = this.ArgToInt(arguments, 6);
        }

        // validate input
        if (settlementDate > maturityDate
            || rate < 0
            || yield < 0
            || redemption <= 0
            || (frequency != 1 && frequency != 2 && frequency != 4)
            || basis < 0
            || basis > 4)
        {
            return this.CreateResult(eErrorType.Num);
        }

        FinanceCalcResult<double>? result = PriceImpl.GetPrice(FinancialDayFactory.Create(settlementDate, (DayCountBasis)basis),
                                                               FinancialDayFactory.Create(maturityDate, (DayCountBasis)basis),
                                                               rate,
                                                               yield,
                                                               redemption,
                                                               frequency,
                                                               (DayCountBasis)basis);

        if (result.HasError)
        {
            return this.CreateResult(result.ExcelErrorType);
        }

        return this.CreateResult(result.Result, DataType.Decimal);
    }
}