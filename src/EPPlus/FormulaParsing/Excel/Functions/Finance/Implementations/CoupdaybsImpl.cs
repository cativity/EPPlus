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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal class CoupdaybsImpl : Coupbase
    {
        public CoupdaybsImpl(FinancialDay settlement, FinancialDay maturity, int frequency, DayCountBasis basis) : base(settlement, maturity, frequency, basis)
        {
        }

        public FinanceCalcResult<int> Coupdaybs()
        {
            IFinanicalDays? fds = FinancialDaysFactory.Create(Basis);
            FinancialPeriod? settlementPeriod = fds.GetCouponPeriod(Settlement, Maturity, Frequency);
            return new FinanceCalcResult<int>((int)Settlement.SubtractDays(settlementPeriod.Start) * -1);
        }
    }
}
