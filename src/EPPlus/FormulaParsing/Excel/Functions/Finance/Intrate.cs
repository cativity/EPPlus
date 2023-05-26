﻿using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "5.5",
        Description = "Calculates the interest rate for a fully invested security")]
    internal class Intrate : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 4);
            System.DateTime settlementDate = System.DateTime.FromOADate(ArgToInt(arguments, 0));
            System.DateTime maturityDate = System.DateTime.FromOADate(ArgToInt(arguments, 1));
            double investment = ArgToDecimal(arguments, 2);
            double redemption = ArgToDecimal(arguments, 3);
            int basis = 0;
            if (arguments.Count() >= 5)
            {
                basis = ArgToInt(arguments, 4);
            }
            if (basis < 0 || basis > 4)
            {
                return this.CreateResult(eErrorType.Num);
            }

            FinanceCalcResult<double>? result = IntRateImpl.Intrate(settlementDate, maturityDate, investment, redemption, (DayCountBasis)basis);
            if (result.HasError)
            {
                return this.CreateResult(result.ExcelErrorType);
            }

            return CreateResult(result.Result, result.DataType);
        }
    }
}
