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
                  Description = "Calculates the cumulative interest paid between two specified periods")]
internal class Cumipmt : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 6);
        double rate = this.ArgToDecimal(arguments, 0);
        int nper = this.ArgToInt(arguments, 1);
        double pv = this.ArgToDecimal(arguments, 2);
        int startPeriod = this.ArgToInt(arguments, 3);
        int endPeriod = this.ArgToInt(arguments, 4);
        int type = this.ArgToInt(arguments, 5);

        if (type < 0 || type > 1)
        {
            return this.CreateResult(eErrorType.Value);
        }

        FinanceCalcResult<double>? result = CumipmtImpl.GetCumipmt(rate, nper, pv, startPeriod, endPeriod, (PmtDue)type);

        if (result.HasError)
        {
            return this.CreateResult(result.ExcelErrorType);
        }

        return this.CreateResult(result.Result, DataType.Decimal);
    }
}