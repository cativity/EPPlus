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

[FunctionMetadata(Category = ExcelFunctionCategory.Text,
                  EPPlusVersion = "4",
                  Description = "Returns the position of a supplied character or text string from within a supplied text string (non-case-sensitive)")]
internal class Search : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        FunctionArgument[]? functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
        ValidateArguments(functionArguments, 2);
        string? search = ArgToString(functionArguments, 0);
        string? searchIn = ArgToString(functionArguments, 1);
        int startIndex = 0;

        if (functionArguments.Count() > 2)
        {
            startIndex = this.ArgToInt(functionArguments, 2) - 1;
        }

        int result = searchIn.IndexOf(search, startIndex, StringComparison.OrdinalIgnoreCase);

        if (result == -1)
        {
            return this.CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
        }

        // Adding 1 because Excel uses 1-based index
        return this.CreateResult(result + 1, DataType.Integer);
    }
}