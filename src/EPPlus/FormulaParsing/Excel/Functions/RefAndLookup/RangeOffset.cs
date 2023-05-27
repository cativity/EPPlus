using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

internal class RangeOffset : ExcelFunction
{
    public IRangeInfo StartRange { get; set; }

    public IRangeInfo EndRange { get; set; }

    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        if (this.StartRange == null || this.EndRange == null)
        {
            return this.CreateResult(eErrorType.Value);
        }

        //Build the address from the minimum row and column to the maximum row and column. StartRange and offsetRange are single cells.
        int fromRow = System.Math.Min(this.StartRange.Address._fromRow, this.EndRange.Address._fromRow);
        int toRow = System.Math.Max(this.StartRange.Address._toRow, this.EndRange.Address._toRow);
        int fromCol = System.Math.Min(this.StartRange.Address._fromCol, this.EndRange.Address._fromCol);
        int toCol = System.Math.Max(this.StartRange.Address._toCol, this.EndRange.Address._toCol);

        EpplusExcelDataProvider.RangeInfo? rangeAddress =
            new EpplusExcelDataProvider.RangeInfo(this.StartRange.Worksheet, new ExcelAddressBase(fromRow, fromCol, toRow, toCol));

        return this.CreateResult(rangeAddress, DataType.Enumerable);
    }
}