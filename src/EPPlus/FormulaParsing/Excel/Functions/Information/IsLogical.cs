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

[FunctionMetadata(Category = ExcelFunctionCategory.Information,
                  EPPlusVersion = "4",
                  Description = "Tests if a supplied value is a logical value, and if so, returns TRUE; Otherwise, returns FALSE")]
internal class IsLogical : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        FunctionArgument[]? functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
        ValidateArguments(functionArguments, 1);
        object? v = GetFirstValue(arguments);

        return this.CreateResult(v is bool, DataType.Boolean);
    }
}