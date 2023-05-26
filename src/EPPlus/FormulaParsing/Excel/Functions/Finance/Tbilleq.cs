﻿/*************************************************************************************************
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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "6.0",
        Description = "Calculates the bond-equivalent yield for a treasury bill")]
    internal class Tbilleq : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            System.DateTime settlementDate = System.DateTime.FromOADate(this.ArgToInt(arguments, 0));
            System.DateTime maturityDate = System.DateTime.FromOADate(this.ArgToInt(arguments, 1));
            double discount = this.ArgToDecimal(arguments, 2);
            if (settlementDate >= maturityDate)
            {
                return this.CreateResult(eErrorType.Num);
            }

            if (discount <= 0d)
            {
                return this.CreateResult(eErrorType.Num);
            }

            IFinanicalDays? finDays = FinancialDaysFactory.Create(DayCountBasis.Actual_360);
            double nDaysInPeriod = finDays.GetDaysBetweenDates(settlementDate, maturityDate);
            if(nDaysInPeriod > 366)
            {
                return this.CreateResult(eErrorType.Num);
            }
            else if(nDaysInPeriod > 182)
            {
                double price = (100d - discount * 100d * nDaysInPeriod / 360d) / 100d;
                int fullYearDays = nDaysInPeriod <= 365 ? 365 : 366;
                double fullYearFactor = nDaysInPeriod / fullYearDays;
                double tmp = System.Math.Pow(fullYearFactor, 2) - (2d * fullYearFactor - 1d) * (1d - 1d / price);
                double term2 = System.Math.Sqrt(tmp);
                double term3 = 2d * fullYearFactor - 1d;
                double result = 2d * (term2 - fullYearFactor) / term3;
                return this.CreateResult(result, DataType.Decimal);
            }
            else
            {
                double result = (365d * discount) / (360d - (discount * nDaysInPeriod));
                return this.CreateResult(result, DataType.Decimal);
            }
        }
    }
}
