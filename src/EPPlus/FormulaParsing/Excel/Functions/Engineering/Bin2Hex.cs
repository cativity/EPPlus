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

[FunctionMetadata(
                     Category = ExcelFunctionCategory.Engineering,
                     EPPlusVersion = "5.1",
                     Description = "Converts a binary number to hexadecimal")]
internal class Bin2Hex : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        string? number = ArgToString(arguments, 0);
        string? formatString = "X";
        if(arguments.Count() > 1)
        {
            int padding = this.ArgToInt(arguments, 1);
            if (padding < 0 ^ padding > 10)
            {
                return this.CreateResult(eErrorType.Num);
            }

            formatString += padding;
        }
        if (number.Length > 10)
        {
            return this.CreateResult(eErrorType.Num);
        }

        if (number.Length < 10)
        {
            int n = Convert.ToInt32(number, 2);
            return this.CreateResult(n.ToString(formatString), DataType.Decimal);
        }
        else
        {
            if (!BinaryHelper.TryParseBinaryToDecimal(number, 2, out int result))
            {
                return this.CreateResult(eErrorType.Num);
            }

            string? hexStr = result.ToString(formatString);
            if(result < 0)
            {
                hexStr = PaddingHelper.EnsureLength(hexStr, 10, "F");
            }
            return this.CreateResult(hexStr, DataType.String);
        }
    }
}