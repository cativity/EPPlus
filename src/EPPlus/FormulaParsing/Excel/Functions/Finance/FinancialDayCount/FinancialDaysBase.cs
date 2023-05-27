/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function
 *************************************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;

internal abstract class FinancialDaysBase
{
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1822:Mark members as static", Justification = "<Pending>")]
    public FinancialPeriod GetCouponPeriod(FinancialDay settlementDay, FinancialDay maturityDay, int frequency)
    {
        //_ = new List<FinancialPeriod>();
        System.DateTime settlementDate = settlementDay.ToDateTime();
        FinancialDay? tmpDay = maturityDay;
        FinancialDay? lastDay = tmpDay;

        while (tmpDay.ToDateTime() > settlementDate)
        {
            switch (frequency)
            {
                case 1:
                    tmpDay = tmpDay.SubtractYears(1);

                    break;

                case 2:
                    tmpDay = tmpDay.SubtractMonths(6, maturityDay.Day);

                    break;

                case 4:
                    tmpDay = tmpDay.SubtractMonths(3, maturityDay.Day);

                    break;

                default:
                    throw new ArgumentException("frequency");
            }

            if (tmpDay > settlementDay)
            {
                lastDay = tmpDay;
            }
        }

        return new FinancialPeriod(tmpDay, lastDay);
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1822:Mark members as static", Justification = "<Pending>")]
    public IEnumerable<FinancialPeriod> GetCouponPeriodsBackwards(FinancialDay settlement, FinancialDay date, int frequency)
    {
        List<FinancialPeriod>? periods = new List<FinancialPeriod>();
        FinancialDay? tmpDay = settlement;

        while (tmpDay > date)
        {
            FinancialDay? periodEndDay = tmpDay;

            switch (frequency)
            {
                case 1:
                    tmpDay = tmpDay.SubtractYears(1);

                    break;

                case 2:
                    tmpDay = tmpDay.SubtractMonths(6, settlement.Day);

                    break;

                case 4:
                    tmpDay = tmpDay.SubtractMonths(3, settlement.Day);

                    break;

                default:
                    throw new ArgumentException("frequency");
            }

            periods.Add(new FinancialPeriod(tmpDay, periodEndDay));
        }

        return periods;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1822:Mark members as static", Justification = "<Pending>")]
    private static FinancialPeriod CreateCalendarPeriod(System.DateTime startDate, int frequency, DayCountBasis basis, bool createFuturePeriod)
    {
        int factor = createFuturePeriod ? 1 : -1;

        System.DateTime d1;
        switch (frequency)
        {
            case 1:
                d1 = startDate.AddYears(1 * factor);

                break;

            case 2:
                d1 = startDate.AddMonths(6 * factor);

                break;

            case 4:
                d1 = startDate.AddMonths(3 * factor);

                break;

            default:
                throw new ArgumentException("frequency");
        }

        if (createFuturePeriod)
        {
            return FinancialDayFactory.CreatePeriod(startDate, d1, basis);
        }
        else
        {
            return FinancialDayFactory.CreatePeriod(d1, startDate, basis);
        }
    }

    private static FinancialPeriod GetSettlementCalendarYearPeriod(FinancialDay date, int frequency)
    {
        System.DateTime startDate;
        if (frequency == 1)
        {
            startDate = new System.DateTime(date.Year, 1, 1);
        }
        else if (frequency == 2)
        {
            if (date.Month < 7)
            {
                startDate = new System.DateTime(date.Year, 1, 1);
            }
            else
            {
                startDate = new System.DateTime(date.Year, 7, 1);
            }
        }
        else if (frequency == 4)
        {
            if (date.Month > 9)
            {
                startDate = new System.DateTime(date.Year, 10, 1);
            }
            else if (date.Month > 6)
            {
                startDate = new System.DateTime(date.Year, 7, 1);
            }
            else if (date.Month > 3)
            {
                startDate = new System.DateTime(date.Year, 4, 1);
            }
            else
            {
                startDate = new System.DateTime(date.Year, 1, 1);
            }
        }
        else
        {
            throw new ArgumentException("frequency");
        }

        return CreateCalendarPeriod(startDate, frequency, date.GetBasis(), true);
    }

    public IEnumerable<FinancialPeriod> GetCalendarYearPeriodsBackwards(FinancialDay settlement, FinancialDay date, int frequency)
    {
        return this.GetCalendarYearPeriodsBackwards(settlement, date, frequency, 0);
    }

    public IEnumerable<FinancialPeriod> GetCalendarYearPeriodsBackwards(FinancialDay settlement, FinancialDay date, int frequency, int additionalPeriods)
    {
        List<FinancialPeriod>? periods = new List<FinancialPeriod>();
        FinancialPeriod? period = GetSettlementCalendarYearPeriod(settlement, frequency);
        periods.Add(period);

        while (period.Start > date)
        {
            System.DateTime dt = period.Start.ToDateTime();
            period = CreateCalendarPeriod(dt, frequency, date.GetBasis(), false);
            periods.Add(period);
        }

        for (int x = 0; x < additionalPeriods; x++)
        {
            System.DateTime tmpDate = periods.Last().Start.ToDateTime();
            FinancialPeriod? p = CreateCalendarPeriod(tmpDate, frequency, date.GetBasis(), false);
            periods.Add(p);
        }

        return periods;
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1822:Mark members as static", Justification = "<Pending>")]
    public int GetNumberOfCouponPeriods(FinancialDay settlementDay, FinancialDay maturityDay, int frequency)
    {
        System.DateTime settlementDate = settlementDay.ToDateTime();
        FinancialDay? tmpDay = maturityDay;
        int nPeriods = 0;

        while (tmpDay.ToDateTime() > settlementDate)
        {
            switch (frequency)
            {
                case 1:
                    tmpDay = tmpDay.SubtractYears(1);

                    break;

                case 2:
                    tmpDay = tmpDay.SubtractMonths(6, maturityDay.Day);

                    break;

                case 4:
                    tmpDay = tmpDay.SubtractMonths(3, maturityDay.Day);

                    break;

                default:
                    throw new ArgumentException("frequency");
            }

            nPeriods++;
        }

        return nPeriods;
    }

    protected virtual double GetDaysBetweenDates(FinancialDay start, FinancialDay end, int basis)
    {
        return (basis * (end.Year - start.Year)) + (30 * (end.Month - start.Month)) + ((end.Day > 30 ? 30 : end.Day) - (start.Day > 30 ? 30 : start.Day));
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Performance", "CA1822:Mark members as static", Justification = "<Pending>")]
    protected double ActualDaysInLeapYear(FinancialDay start, FinancialDay end)
    {
        double daysInLeapYear = 0d;

        for (short year = start.Year; year <= end.Year; year++)
        {
            if (System.DateTime.IsLeapYear(year))
            {
                if (year == start.Year)
                {
                    daysInLeapYear += new System.DateTime(year + 1, 1, 1).Subtract(start.ToDateTime()).TotalDays;
                }
                else if (year == end.Year)
                {
                    daysInLeapYear += end.ToDateTime().Subtract(new System.DateTime(year, 1, 1)).TotalDays;
                }
                else
                {
                    daysInLeapYear += 366d;
                }
            }
        }

        return daysInLeapYear;
    }
}