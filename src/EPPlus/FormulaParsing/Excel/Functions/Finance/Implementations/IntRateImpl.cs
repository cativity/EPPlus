using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;

internal class IntRateImpl
{
    internal static FinanceCalcResult<double> Intrate(System.DateTime settlement,
                                                      System.DateTime maturity,
                                                      double investment,
                                                      double redemption,
                                                      DayCountBasis basis = DayCountBasis.US_30_360)
    {
        if (investment <= 0 || redemption <= 0)
        {
            return new FinanceCalcResult<double>(eErrorType.Num);
        }

        if (maturity <= settlement)
        {
            return new FinanceCalcResult<double>(eErrorType.Num);
        }

        FinancialDay? settlementDay = FinancialDayFactory.Create(settlement, basis);
        FinancialDay? maturityDay = FinancialDayFactory.Create(maturity, basis);
        IFinanicalDays? fd = FinancialDaysFactory.Create(basis);
        double nDays = fd.GetDaysBetweenDates(settlementDay, maturityDay);

        // special case to make this function return same value as Excel
        if (basis == DayCountBasis.US_30_360 && maturityDay.Day == 31)
        {
            nDays++;
        }

        double result = (redemption - investment) / investment * fd.DaysPerYear / nDays;

        return new FinanceCalcResult<double>(result);
    }
}