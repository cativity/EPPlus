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

[FunctionMetadata(Category = ExcelFunctionCategory.LookupAndReference,
                  EPPlusVersion = "4",
                  Description = "Returns a cell or range reference that is represented by a supplied text string")]
internal class Indirect : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        string? address = ArgToAddress(arguments, 0);
        ExcelAddressBase adr;

        if (ExcelCellBase.IsValidAddress(address) || ExcelAddressBase.IsTableAddress(address))
        {
            adr = new ExcelAddressBase(address);
        }
        else
        {
            INameInfo? n = context.ExcelDataProvider.GetName(context.Scopes.Current.Address.Worksheet, address);

            if (n.Value is EpplusExcelDataProvider.RangeInfo ri)
            {
                adr = ri.Address;
            }
            else
            {
                adr = new ExcelAddressBase(n.Formula);
            }

            address = adr.Address;
        }

        string? ws = adr.WorkSheetName;

        if (string.IsNullOrEmpty(ws))
        {
            ws = context.Scopes.Current.Address.Worksheet;
        }

        IRangeInfo? result = context.ExcelDataProvider.GetRange(ws, context.Scopes.Current.Address.FromRow, context.Scopes.Current.Address.FromCol, address);

        if (result.IsEmpty)
        {
            return CompileResult.Empty;
        }
        else if (!result.IsMulti)
        {
            ICellInfo? cell = result.FirstOrDefault();
            object? val = cell?.Value;

            if (val == null)
            {
                return CompileResult.Empty;
            }

            return new CompileResultFactory().Create(val);
        }

        return new CompileResult(result, DataType.Enumerable);
    }
}