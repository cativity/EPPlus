/*************************************************************************************************
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

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

[FunctionMetadata(Category = ExcelFunctionCategory.Text, EPPlusVersion = "4", Description = "Joins together two or more text strings")]
internal class Concatenate : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        if (arguments == null)
        {
            return this.CreateResult(string.Empty, DataType.String);
        }

        StringBuilder? sb = new StringBuilder();

        foreach (FunctionArgument? arg in arguments)
        {
            object? v = arg.ValueFirst;

            if (v != null)
            {
                _ = sb.Append(v);
            }
        }

        return this.CreateResult(sb.ToString(), DataType.String);
    }
}