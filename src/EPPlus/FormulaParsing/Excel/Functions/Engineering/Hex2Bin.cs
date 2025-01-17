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

using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering;

[FunctionMetadata(Category = ExcelFunctionCategory.Engineering, EPPlusVersion = "5.1", Description = "Converts a hexadecimal number to binary")]
internal class Hex2Bin : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        string? number = ArgToString(arguments, 0);
        int? padding = default(int?);

        if (arguments.Count() > 1)
        {
            padding = this.ArgToInt(arguments, 1);

            if ((padding.Value < 0) ^ (padding.Value > 10))
            {
                return this.CreateResult(eErrorType.Num);
            }
        }

        double decNumber = TwoComplementHelper.ParseDecFromString(number, 16);
        string? result = Convert.ToString((int)decNumber, 2);

        if (padding.HasValue)
        {
            result = PaddingHelper.EnsureLength(result, padding.Value, "0");
        }
        else
        {
            result = PaddingHelper.EnsureMinLength(result, 10);
        }

        return this.CreateResult(result, DataType.String);
    }
}