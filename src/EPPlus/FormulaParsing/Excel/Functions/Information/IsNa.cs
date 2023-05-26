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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

[FunctionMetadata(
                     Category = ExcelFunctionCategory.Information,
                     EPPlusVersion = "4",
                     Description = "Tests if an initial supplied value (or expression) returns the Excel #N/A error and if so, returns TRUE; Otherwise returns FALSE")]
internal class IsNa : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        if (arguments == null || arguments.Count() == 0)
        {
            return this.CreateResult(false, DataType.Boolean);
        }

        object? v = GetFirstValue(arguments);

        if (v is ExcelErrorValue && ((ExcelErrorValue)v).Type==eErrorType.NA)
        {
            return this.CreateResult(true, DataType.Boolean);
        }
        return this.CreateResult(false, DataType.Boolean);
    }
}