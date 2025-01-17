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

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering;

[FunctionMetadata(Category = ExcelFunctionCategory.Engineering, EPPlusVersion = "5.1", Description = "Tests whether two supplied numbers are equal")]
internal class Delta : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        double n1 = this.ArgToDecimal(arguments, 0);
        double n2 = 0d;

        if (arguments.Count() > 1)
        {
            n2 = this.ArgToDecimal(arguments, 1);
        }

        if (n1.CompareTo(n2) == 0)
        {
            return this.CreateResult(1, DataType.Integer);
        }

        return this.CreateResult(0, DataType.Integer);
    }
}