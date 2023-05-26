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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

[FunctionMetadata(
                     Category = ExcelFunctionCategory.LookupAndReference,
                     EPPlusVersion = "4",
                     Description = "Returns a reference to a range of cells that is a specified number of rows and columns from an initial supplied range")]
internal class Offset : LookupFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        FunctionArgument[]? functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
        ValidateArguments(functionArguments, 3);
        string? startRange = ArgToAddress(functionArguments, 0, context);
        int rowOffset = this.ArgToInt(functionArguments, 1);
        int colOffset = this.ArgToInt(functionArguments, 2);
        int width = 0, height = 0;
        if (functionArguments.Length > 3)
        {
            height = this.ArgToInt(functionArguments, 3);
            if (height == 0)
            {
                return new CompileResult(eErrorType.Ref);
            }
        }
        if (functionArguments.Length > 4)
        {
            width = this.ArgToInt(functionArguments, 4);
            if (width == 0)
            {
                return new CompileResult(eErrorType.Ref);
            }
        }
        string? ws = context.Scopes.Current.Address.Worksheet;            
        IRangeInfo? r =context.ExcelDataProvider.GetRange(ws, context.Scopes.Current.Address.FromRow, context.Scopes.Current.Address.FromCol, startRange);
        ExcelAddressBase? adr = r.Address;

        int fromRow = adr._fromRow + rowOffset;
        int fromCol = adr._fromCol + colOffset;
        int toRow = (height != 0 ? adr._fromRow + height - 1 : adr._toRow) + rowOffset;
        int toCol = (width != 0 ? adr._fromCol + width - 1 : adr._toCol) + colOffset;

        IRangeInfo? newRange = context.ExcelDataProvider.GetRange(adr.WorkSheetName, fromRow, fromCol, toRow, toCol);
            
        return this.CreateResult(newRange, DataType.Enumerable);
    }
}