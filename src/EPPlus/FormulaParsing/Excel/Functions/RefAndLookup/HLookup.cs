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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

[FunctionMetadata(Category = ExcelFunctionCategory.LookupAndReference,
                  EPPlusVersion = "4",
                  Description = "Looks up a supplied value in the first row of a table, and returns the corresponding value from another row")]
internal class HLookup : LookupFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 3);
        LookupArguments? lookupArgs = new LookupArguments(arguments, context);

        if (lookupArgs.LookupIndex < 1)
        {
            return this.CreateResult(eErrorType.Value);
        }

        LookupNavigator? navigator = LookupNavigatorFactory.Create(LookupDirection.Horizontal, lookupArgs, context);

        return this.Lookup(navigator, lookupArgs);
    }
}