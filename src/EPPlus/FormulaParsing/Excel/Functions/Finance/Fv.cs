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
                  Description = "Calculates the future value of an investment with periodic constant payments and a constant interest rate")]
internal class Fv : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 2);
        double rate = this.ArgToDecimal(arguments, 0);
        double nPer = this.ArgToDecimal(arguments, 1);
        double pmt = 0d;

        if (arguments.Count() >= 3)
        {
            pmt = this.ArgToDecimal(arguments, 2);
        }

        double pv = 0d;

        if (arguments.Count() >= 4)
        {
            pv = this.ArgToDecimal(arguments, 3);
        }

        int type = 0;

        if (arguments.Count() >= 5)
        {
            type = this.ArgToInt(arguments, 4);
        }

        FinanceCalcResult<double>? retVal = FvImpl.Fv(rate, nPer, pmt, pv, (PmtDue)type);

        if (retVal.HasError)
        {
            return this.CreateResult(retVal.ExcelErrorType);
        }

        return this.CreateResult(retVal.Result, DataType.Decimal);
    }
}