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

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

[FunctionMetadata(Category = ExcelFunctionCategory.Text,
                  EPPlusVersion = "5.0",
                  Description = "Removes all non-printable characters from a supplied text string")]
internal class Clean : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        string? str = ArgToString(arguments, 0);

        if (!string.IsNullOrEmpty(str))
        {
            StringBuilder? sb = new StringBuilder();
            byte[]? arr = Encoding.ASCII.GetBytes(str);

            foreach (byte c in arr)
            {
                if (c > 31)
                {
                    _ = sb.Append((char)c);
                }
            }

            str = sb.ToString();
        }

        return this.CreateResult(str, DataType.String);
    }
}