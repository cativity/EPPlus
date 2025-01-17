﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/13/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;

[FunctionMetadata(Category = ExcelFunctionCategory.Financial,
                  EPPlusVersion = "5.2",
                  Description = "Calculates the number of days in the coupon period that contains the settlement date")]
internal class Coupdays : CoupFunctionBase<double>
{
    protected override FinanceCalcResult<double> ExecuteFunction(FinancialDay settlementDate,
                                                                 FinancialDay maturityDate,
                                                                 int frequency,
                                                                 DayCountBasis basis = DayCountBasis.US_30_360)
    {
        CoupdaysImpl? impl = new CoupdaysImpl(settlementDate, maturityDate, frequency, basis);

        return impl.GetCoupdays();
    }
}