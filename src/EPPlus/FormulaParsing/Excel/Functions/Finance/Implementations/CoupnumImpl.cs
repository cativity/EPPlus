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
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;

internal class CoupnumImpl : Coupbase
{
    public CoupnumImpl(FinancialDay settlement, FinancialDay maturity, int frequency, DayCountBasis basis)
        : base(settlement, maturity, frequency, basis)
    {
    }

    internal FinanceCalcResult<int> GetCoupnum()
    {
        IFinanicalDays? fds = FinancialDaysFactory.Create(this.Basis);

        return new FinanceCalcResult<int>(fds.GetNumberOfCouponPeriods(this.Settlement, this.Maturity, this.Frequency));
    }
}