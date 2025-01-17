﻿/*************************************************************************************************
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

/// <summary>
/// Rules as defined on https://en.wikipedia.org/wiki/Day_count_convention
/// </summary>
internal class FinancialDaysUs_30_360 : FinancialDaysBase, IFinanicalDays
{
    public double GetDaysBetweenDates(System.DateTime startDate, System.DateTime endDate)
    {
        FinancialDay? start = FinancialDayFactory.Create(startDate, DayCountBasis.US_30_360);
        FinancialDay? end = FinancialDayFactory.Create(endDate, DayCountBasis.US_30_360);

        return this.GetDaysBetweenDates(start, end);
    }

    public double GetDaysBetweenDates(FinancialDay startDate, FinancialDay endDate)
    {
        if (endDate.IsLastDayOfFebruary)
        {
            if (startDate.IsLastDayOfFebruary)
            {
                endDate.Day = 30;
                startDate.Day = 30;
            }
            else
            {
                endDate.Day = 30;
            }
        }

        if (endDate.Day == 31 && (startDate.Day == 30 || startDate.Day == 31))
        {
            endDate.Day = 30;
        }

        if (startDate.Day == 31)
        {
            startDate.Day = 30;
        }

        return this.GetDaysBetweenDates(startDate, endDate, (int)this.DaysPerYear);
    }

    public double GetCoupdays(FinancialDay start, FinancialDay end, int frequency) => this.DaysPerYear / frequency;

    public double DaysPerYear => 360d;
}