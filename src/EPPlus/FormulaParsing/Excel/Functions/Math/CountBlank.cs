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

using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

[FunctionMetadata(Category = ExcelFunctionCategory.Statistical, EPPlusVersion = "4", Description = "Returns the number of blank cells in a supplied range")]
internal class CountBlank : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        FunctionArgument? arg = arguments.First();

        if (!arg.IsExcelRange && arg.ExcelAddressReferenceId <= 0)
        {
            throw new InvalidOperationException("CountBlank only support ranges as arguments");
        }

        IRangeInfo range;


        int result;
        if (arg.IsExcelRange)
        {
            range = arg.ValueAsRangeInfo;
            result = arg.ValueAsRangeInfo.GetNCells();
        }
        else
        {
            RangeAddress? currentCell = context.Scopes.Current.Address;
            string? worksheet = currentCell.Worksheet;
            string? address = context.AddressCache.Get(arg.ExcelAddressReferenceId);
            ExcelAddressBase? excelAddress = new ExcelAddressBase(address);

            if (!string.IsNullOrEmpty(excelAddress.WorkSheetName))
            {
                worksheet = excelAddress.WorkSheetName;
            }

            range = context.ExcelDataProvider.GetRange(worksheet, currentCell.FromRow, currentCell.FromCol, excelAddress.Address);
            result = range.GetNCells();
        }

        foreach (ICellInfo? cell in range)
        {
            if (!(cell.Value == null || cell.Value.ToString() == string.Empty))
            {
                result--;
            }
        }

        return this.CreateResult(result, DataType.Integer);
    }
}