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
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;

internal abstract class FinancialDay
{
    public FinancialDay(System.DateTime date)
    {
        this.Year = Convert.ToInt16(date.Year);
        this.Month = Convert.ToInt16(date.Month);
        this.Day = Convert.ToInt16(date.Day);
    }

    public FinancialDay(int year, int month, int day)
    {
        this.Year = (short)year;
        this.Month = (short)month;
        this.Day = (short)day;
    }

    public override string ToString()
    {
        return $"{this.Year}-{this.Month}-{this.Day}";
    }

    public short Year { get; set; }

    public short Month { get; set; }

    public short Day { get; set; }

    public bool IsLastDayOfFebruary
    {
        get { return this.Month == 2 && this.Day == System.DateTime.DaysInMonth(this.Year, this.Month); }
    }

    public bool IsLastDayOfMonth
    {
        get { return this.Day == System.DateTime.DaysInMonth(this.Year, this.Month); }
    }

    public System.DateTime ToDateTime()
    {
        return new System.DateTime(this.Year, this.Month, this.Day);
    }

    public FinancialDay SubtractYears(int years)
    {
        short day = this.Day;

        if (this.IsLastDayOfFebruary && System.DateTime.IsLeapYear(this.Year) && !System.DateTime.IsLeapYear(this.Year + years))
        {
            day -= 1;
        }

        return this.Factory((short)(this.Year - years), this.Month, day);
    }

    public int CompareTo(FinancialDay other)
    {
        if (other is null)
        {
            return 1;
        }

        if (this.Year == other.Year && this.Month == other.Month && this.Day == other.Day)
        {
            return 0;
        }

        return this.ToDateTime().CompareTo(other.ToDateTime());
    }

    public static bool operator >(FinancialDay a, FinancialDay b) => a.CompareTo(b) > 0;

    public static bool operator <(FinancialDay a, FinancialDay b) => a.CompareTo(b) < 0;

    public static bool operator <=(FinancialDay a, FinancialDay b) => a.CompareTo(b) <= 0;

    public static bool operator >=(FinancialDay a, FinancialDay b) => a.CompareTo(b) >= 0;

    public static bool operator ==(FinancialDay a, FinancialDay b)
    {
        if (a is null && b is null)
        {
            return true;
        }

        if (!(a is null) && b is null)
        {
            return false;
        }

        if (a is null && !(b is null))
        {
            return false;
        }

        return a.CompareTo(b) == 0;
    }

    public static bool operator !=(FinancialDay a, FinancialDay b)
    {
        if (a is null && b is null)
        {
            return false;
        }

        if (!(a is null) && b is null)
        {
            return true;
        }

        if (a is null && !(b is null))
        {
            return true;
        }

        return a.CompareTo(b) != 0;
    }

    public FinancialDay SubtractMonths(int months, short day)
    {
        short year = this.Year;
        short actualDay = day;
        short month = this.Month;

        if (this.Month - months < 1)
        {
            year -= 1;
            month = Convert.ToInt16(12 - ((this.Month - months) * -1));
        }
        else
        {
            month = (short)(this.Month - Convert.ToInt16(months));
        }

        if (this.IsLastDayOfFebruary && System.DateTime.IsLeapYear(this.Year) && !System.DateTime.IsLeapYear(year))
        {
            actualDay -= 1;
        }
        else if (System.DateTime.DaysInMonth(year, month) < actualDay)
        {
            actualDay = (short)System.DateTime.DaysInMonth(year, month);
        }

        return this.Factory(year, month, actualDay);
    }

    protected abstract FinancialDay Factory(short year, short month, short day);

    internal DayCountBasis GetBasis()
    {
        return this.Basis;
    }

    protected abstract DayCountBasis Basis { get; }

    /// <summary>
    /// Number of days between two <see cref="FinancialDay"/>s
    /// </summary>
    /// <param name="day">The other day</param>
    /// <returns>Number of days according to the <see cref="DayCountBasis"/> of this day</returns>
    public double SubtractDays(FinancialDay day)
    {
        IFinanicalDays? financialDays = FinancialDaysFactory.Create(this.Basis);

        return financialDays.GetDaysBetweenDates(this.ToDateTime(), day.ToDateTime());
    }

    public override bool Equals(object obj)
    {
        if (obj is FinancialDay b)
        {
            return b == this;
        }
        else
        {
            return false;
        }
    }

    public override int GetHashCode()
    {
        return base.GetHashCode();
    }
}