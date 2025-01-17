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
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

[FunctionMetadata(Category = ExcelFunctionCategory.Text,
                  EPPlusVersion = "5.0",
                  Description = "Joins together two or more text strings",
                  IntroducedInExcelVersion = "2016")]
internal class Concat : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        if (arguments == null)
        {
            return this.CreateResult(string.Empty, DataType.String);
        }
        else if (arguments.Count() > 254)
        {
            return this.CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelAddress);
        }

        StringBuilder? sb = new StringBuilder();

        foreach (FunctionArgument? arg in arguments)
        {
            if (arg.IsExcelRange)
            {
                foreach (ICellInfo? cell in arg.ValueAsRangeInfo)
                {
                    _ = sb.Append(cell.Value);
                }
            }
            else
            {
                object? v = arg.ValueFirst;

                if (v != null)
                {
                    _ = sb.Append(v);
                }
            }
        }

        string? result = sb.ToString();

        if (!string.IsNullOrEmpty(result) && result.Length > 32767)
        {
            return this.CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelAddress);
        }

        return this.CreateResult(sb.ToString(), DataType.String);
    }
}