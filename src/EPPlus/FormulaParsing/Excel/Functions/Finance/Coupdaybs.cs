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
                  Description = "Calculates the number of days from the beginning of the coupon period to the settlement date")]
internal class Coupdaybs : CoupFunctionBase<int>
{
    protected override FinanceCalcResult<int> ExecuteFunction(FinancialDay settlementDate,
                                                              FinancialDay maturityDate,
                                                              int frequency,
                                                              DayCountBasis basis = DayCountBasis.US_30_360)
    {
        CoupdaybsImpl? impl = new CoupdaybsImpl(settlementDate, maturityDate, frequency, basis);

        return impl.Coupdaybs();
    }
}