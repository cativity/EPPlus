using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.Table;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information;

[FunctionMetadata(Category = ExcelFunctionCategory.Information,
                  EPPlusVersion = "5.5",
                  IntroducedInExcelVersion = "2013",
                  Description = "Returns the sheet number relating to a supplied reference")]
internal class Sheet : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        int result = -1;

        if (arguments.Count() == 0)
        {
            RangeAddress? cell = context.Scopes.Current.Address;
            string? ws = cell.Worksheet;
            result = context.ExcelDataProvider.GetWorksheetIndex(ws);
        }
        else
        {
            FunctionArgument? arg = arguments.ElementAt(0);

            if (arg.ExcelAddressReferenceId > 0)
            {
                string? address = ArgToAddress(arguments, 0, context);

                if (address.Contains('!'))
                {
                    ExcelAddress? excelAddress = new ExcelAddress(address);
                    result = context.ExcelDataProvider.GetWorksheetIndex(excelAddress.WorkSheetName);
                }
                else
                {
                    string? value = string.IsNullOrEmpty(address) ? ArgToString(arguments, 0) : address;
                    IEnumerable<string>? worksheetNames = context.ExcelDataProvider.GetWorksheets();

                    // for each worksheet in the workbook - check if the value a worksheet name.
                    foreach (string? wsName in worksheetNames)
                    {
                        if (string.Compare(wsName, value, true) == 0)
                        {
                            result = context.ExcelDataProvider.GetWorksheetIndex(wsName);

                            break;
                        }
                    }

                    if (result == -1)
                    {
                        // not a worksheet name, now check if it is a named range in the current worksheet
                        ExcelNamedRangeCollection? wsNamedRanges = context.ExcelDataProvider.GetWorksheetNames(context.Scopes.Current.Address.Worksheet);
                        ExcelNamedRange? matchingWsName = wsNamedRanges.FirstOrDefault(x => x.Name == value);

                        if (matchingWsName != null)
                        {
                            result = context.ExcelDataProvider.GetWorksheetIndex(matchingWsName.WorkSheetName);
                        }

                        if (result == -1)
                        {
                            // not a worksheet named range, now check workbook level
                            ExcelNamedRangeCollection? namedRanges = context.ExcelDataProvider.GetWorkbookNameValues();
                            ExcelNamedRange? matchingWorkbookRange = namedRanges.FirstOrDefault(x => x.Name == value);

                            if (matchingWorkbookRange != null)
                            {
                                result = context.ExcelDataProvider.GetWorksheetIndex(matchingWorkbookRange.WorkSheetName);
                            }
                            else
                            {
                                result = context.ExcelDataProvider.GetWorksheetIndex(value);
                            }
                        }

                        if (result == -1)
                        {
                            ExcelTable? table = context.ExcelDataProvider.GetExcelTable(value);

                            if (table != null)
                            {
                                result = context.ExcelDataProvider.GetWorksheetIndex(table.WorkSheet.Name);
                            }
                        }
                    }
                }
            }
            else
            {
                string? value = ArgToString(arguments, 0);
                result = context.ExcelDataProvider.GetWorksheetIndex(value);
            }
        }

        if (result == -1)
        {
            return this.CreateResult(eErrorType.NA);
        }

        return this.CreateResult(result, DataType.Integer);
    }
}