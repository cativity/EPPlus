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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Financial,
        EPPlusVersion = "6.0",
        Description = "Calculates the accrued interest for a security that pays periodic interest.")]
    internal class Accrint : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 6);
            // collect input
            System.DateTime issueDate = System.DateTime.FromOADate(ArgToInt(arguments, 0));
            System.DateTime firstInterestDate = System.DateTime.FromOADate(ArgToInt(arguments, 1));
            System.DateTime settlementDate = System.DateTime.FromOADate(ArgToInt(arguments, 2));
            double rate = ArgToDecimal(arguments, 3);
            double par = ArgToDecimal(arguments, 4);
            int frequency = ArgToInt(arguments, 5);
            int basis = 0;
            if(arguments.Count() >= 7)
            {
                basis = ArgToInt(arguments, 6);
            }
            bool issueToSettlement = true;
            if(arguments.Count() >= 8)
            {
                issueToSettlement = ArgToBool(arguments, 7);
            }

            // validate input
            if (rate <= 0 || par <= 0)
            {
                return this.CreateResult(eErrorType.Num);
            }

            if (frequency != 1 && frequency != 2 && frequency != 4)
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

            // calculation
            DayCountBasis dayCountBasis = (DayCountBasis)basis;
            IFinanicalDays? financialDays = FinancialDaysFactory.Create(dayCountBasis);
            FinancialDay? issue = FinancialDayFactory.Create(issueDate, dayCountBasis);
            FinancialDay? settlement = FinancialDayFactory.Create(settlementDate, dayCountBasis);
            FinancialDay? firstInterest = FinancialDayFactory.Create(firstInterestDate.AddDays(firstInterestDate.Day * -1 + 1), dayCountBasis);
            
            if(issueToSettlement)
            {
                YearFracProvider? yearFrac = new YearFracProvider(context);
                double r = yearFrac.GetYearFrac(issueDate, settlementDate, dayCountBasis) * rate * par;
                return CreateResult(r, DataType.Decimal);
            }
            else
            {
                double r = CalculateInterest(issue, firstInterest, settlement, rate, par, frequency, dayCountBasis, context);
                return CreateResult(r, DataType.Decimal);
            }
        }

        private static double CalculateInterest(FinancialDay issue, FinancialDay firstInterest, FinancialDay settlement, double rate, double par, int frequency, DayCountBasis basis, ParsingContext context)
        {
            YearFracProvider? yearFrac = new YearFracProvider(context);
            IFinanicalDays? fds = FinancialDaysFactory.Create(basis);
            int nAdditionalPeriods = frequency == 1 ? 0 : 1;
            if(firstInterest <= settlement)
            {
                IEnumerable<FinancialPeriod>? p = fds.GetCalendarYearPeriodsBackwards(settlement, firstInterest, frequency, nAdditionalPeriods);
                IEnumerable<FinancialPeriod>? p2 = fds.GetCalendarYearPeriodsBackwards(firstInterest, settlement, frequency, nAdditionalPeriods);
                FinancialPeriod? firstPeriod = settlement >= firstInterest ? p.Last() : p.First();
                double yearFrac2 = yearFrac.GetYearFrac(firstPeriod.Start.ToDateTime(), settlement.ToDateTime(), basis);
                return yearFrac2 * rate * par;
            }
            else
            {
                IEnumerable<FinancialPeriod>? p2 = fds.GetCalendarYearPeriodsBackwards(firstInterest, settlement, frequency, nAdditionalPeriods);
                FinancialPeriod? firstInterestPeriod = p2.FirstOrDefault(x => x.Start < firstInterest && x.End >= firstInterest);
                double yearFrac2 = yearFrac.GetYearFrac(settlement.ToDateTime(), firstInterestPeriod.Start.ToDateTime(), basis) * -1;
                return yearFrac2 * rate * par;
            }
        }
    }
}
