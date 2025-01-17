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

internal class FinancialDays_Actual_360 : FinancialDaysBase, IFinanicalDays
{
    public double GetDaysBetweenDates(System.DateTime startDate, System.DateTime endDate)
    {
        FinancialDay? start = FinancialDayFactory.Create(startDate, DayCountBasis.Actual_360);
        FinancialDay? end = FinancialDayFactory.Create(endDate, DayCountBasis.Actual_360);

        return this.GetDaysBetweenDates(start, end, 360);
    }

    public double GetDaysBetweenDates(FinancialDay startDate, FinancialDay endDate) => this.GetDaysBetweenDates(startDate, endDate, (int)this.DaysPerYear);

    protected override double GetDaysBetweenDates(FinancialDay start, FinancialDay end, int basis) => end.ToDateTime().Subtract(start.ToDateTime()).TotalDays;

    public double GetCoupdays(FinancialDay start, FinancialDay end, int frequency) => this.DaysPerYear / frequency;

    public double DaysPerYear => 360d;
}