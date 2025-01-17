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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;

[FunctionMetadata(Category = ExcelFunctionCategory.Logical,
                  EPPlusVersion = "4",
                  Description = "Tests a user-defined condition and returns one result if the condition is TRUE, and another result if the condition is FALSE")]
internal class If : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 3);
        bool condition = this.ArgToBool(arguments, 0);
        object? firstStatement = arguments.ElementAt(1).Value;
        object? secondStatement = arguments.ElementAt(2).Value;
        CompileResultFactory? factory = new CompileResultFactory();

        return condition ? factory.Create(firstStatement) : factory.Create(secondStatement);
    }
}