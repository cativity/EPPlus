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
                  Description = "Calculates the number of periods required for an investment to reach a specified value",
                  IntroducedInExcelVersion = "2013")]
internal class Pduration : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 3);
        double rate = this.ArgToDecimal(arguments, 0);
        double pv = this.ArgToDecimal(arguments, 1);
        double fv = this.ArgToDecimal(arguments, 2);

        if (rate <= 0d || pv <= 0d || fv <= 0d)
        {
            return this.CreateResult(eErrorType.Num);
        }

        double retVal = (System.Math.Log(fv) - System.Math.Log(pv)) / System.Math.Log(1 + rate);

        return this.CreateResult(retVal, DataType.Decimal);
    }
}