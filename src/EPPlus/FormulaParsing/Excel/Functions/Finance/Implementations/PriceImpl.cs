﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;

internal static class PriceImpl
{
    public static FinanceCalcResult<double> GetPrice(FinancialDay settlement, FinancialDay maturity, double rate, double yield, double redemption, int frequency, DayCountBasis basis = DayCountBasis.US_30_360)
    {
        FinanceCalcResult<double>? coupDaysResult = new CoupdaysImpl(settlement, maturity, frequency, basis).GetCoupdays();
        if (coupDaysResult.HasError)
        {
            return coupDaysResult;
        }

        FinanceCalcResult<double>? coupdaysNcResult = new CoupdaysncImpl(settlement, maturity, frequency, basis).Coupdaysnc();
        if (coupdaysNcResult.HasError)
        {
            return coupdaysNcResult;
        }

        FinanceCalcResult<int>? coupnumResult = new CoupnumImpl(settlement, maturity, frequency, basis).GetCoupnum();
        if (coupnumResult.HasError)
        {
            return new FinanceCalcResult<double>(coupnumResult.ExcelErrorType);
        }

        FinanceCalcResult<int>? coupdaysbsResult = new CoupdaybsImpl(settlement, maturity, frequency, basis).Coupdaybs();
        if(coupdaysbsResult.HasError)
        {
            return new FinanceCalcResult<double>(coupdaysbsResult.ExcelErrorType);
        }

        double E = coupDaysResult.Result;
        double DSC = coupdaysNcResult.Result;
        int N = coupnumResult.Result;
        int A = coupdaysbsResult.Result;

        double retVal = -1d;
        if(N > 1)
        {
            double part1 = redemption / System.Math.Pow(1d + (yield / frequency), N - 1d + (DSC / E));
            double sum = 0d;
            for (int k = 1; k <= N; k++)
            {
                sum += (100 * (rate / frequency)) / System.Math.Pow(1 + yield / frequency, k - 1 + DSC / E);
            }

            retVal = part1 + sum - (100 * (rate / frequency) * (A / E));
        }
        else
        {
            double DSR = E - A;
            double T1 = 100 * (rate / frequency) + redemption;
            double T2 = (yield / frequency) * (DSR / E) + 1;
            double T3 = 100 * (rate / frequency) * (A / E);

            retVal = T1 / T2 - T3;
        }

        return new FinanceCalcResult<double>(retVal);
    }
}