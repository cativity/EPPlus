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

[FunctionMetadata(Category = ExcelFunctionCategory.LookupAndReference, EPPlusVersion = "4", Description = "Returns the number of rows in a supplied range")]
internal class Rows : LookupFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        IRangeInfo? r = arguments.ElementAt(0).ValueAsRangeInfo;

        if (r != null)
        {
            return this.CreateResult(r.Address._toRow - r.Address._fromRow + 1, DataType.Integer);
        }
        else
        {
            string? range = ArgToAddress(arguments, 0, context);

            if (ExcelAddressUtil.IsValidAddress(range))
            {
                RangeAddressFactory? factory = new RangeAddressFactory(context.ExcelDataProvider);
                RangeAddress? address = factory.Create(range);

                return this.CreateResult(address.ToRow - address.FromRow + 1, DataType.Integer);
            }
        }

        throw new ArgumentException("Invalid range supplied");
    }
}