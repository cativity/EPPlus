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

[FunctionMetadata(
                     Category = ExcelFunctionCategory.Text,
                     EPPlusVersion = "4",
                     Description = "Tests if two supplied text strings are exactly the same and if so, returns TRUE; Otherwise, returns FALSE. (case-sensitive)")]
internal class Exact : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 2);
        object? val1 = arguments.ElementAt(0).ValueFirst;
        object? val2 = arguments.ElementAt(1).ValueFirst;

        if (val1 == null && val2 == null)
        {
            return this.CreateResult(true, DataType.Boolean);
        }
        else if ((val1 == null && val2 != null) || (val1 != null && val2 == null))
        {
            return this.CreateResult(false, DataType.Boolean);
        }

        int result = string.Compare(val1.ToString(), val2.ToString(), StringComparison.Ordinal);
        return this.CreateResult(result == 0, DataType.Boolean);
    }
}