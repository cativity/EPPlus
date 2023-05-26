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
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

[FunctionMetadata(
                     Category = ExcelFunctionCategory.LookupAndReference,
                     EPPlusVersion = "4",
                     Description = "Returns the row number of a supplied range, or of the current cell")]
internal class Row : LookupFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        if (arguments == null || arguments.Count() == 0)
        {
            return this.CreateResult(context.Scopes.Current.Address.FromRow, DataType.Integer);
        }
        string? rangeAddress = ArgToAddress(arguments, 0, context);
        if (!ExcelAddressUtil.IsValidAddress(rangeAddress))
        {
            throw new ArgumentException("An invalid argument was supplied");
        }

        RangeAddressFactory? factory = new RangeAddressFactory(context.ExcelDataProvider);
        RangeAddress? address = factory.Create(rangeAddress);
        return this.CreateResult(address.FromRow, DataType.Integer);
    }
}