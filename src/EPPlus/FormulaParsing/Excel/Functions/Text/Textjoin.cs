/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  06/18/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

[FunctionMetadata(
                     Category = ExcelFunctionCategory.Text,
                     EPPlusVersion = "5.2",
                     IntroducedInExcelVersion = "2019",
                     Description = "Joins together two or more text strings, separated by a delimiter")]
internal class Textjoin : ExcelFunction
{
    private readonly int MaxReturnLength = 32767;

    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 3);
        string? delimiter = ArgToString(arguments, 0);
        bool ignoreEmpty = this.ArgToBool(arguments, 1);
        StringBuilder? str = new StringBuilder();
        for(int x = 2; x < arguments.Count() && x < 252; x++)
        {
            FunctionArgument? arg = arguments.ElementAt(x);
            if(arg.IsExcelRange)
            {
                foreach(ICellInfo? cell in arg.ValueAsRangeInfo)
                {
                    string? val = cell.Value != null ? cell.Value.ToString() : string.Empty;
                    if (ignoreEmpty && string.IsNullOrEmpty(val))
                    {
                        continue;
                    }

                    str.Append(val);
                    str.Append(delimiter);
                    if (str.Length > this.MaxReturnLength)
                    {
                        return this.CreateResult(eErrorType.Value);
                    }
                }
            }
            else if(arg.Value is IEnumerable<FunctionArgument>)
            {
                IEnumerable<FunctionArgument>? items = arg.Value as IEnumerable<FunctionArgument>;
                if(items != null)
                {
                    foreach(FunctionArgument? item in items)
                    {
                        string? val = item.Value != null ? item.Value.ToString() : string.Empty;
                        if (ignoreEmpty && string.IsNullOrEmpty(val))
                        {
                            continue;
                        }

                        str.Append(val);
                        str.Append(delimiter);
                        if (str.Length > this.MaxReturnLength)
                        {
                            return this.CreateResult(eErrorType.Value);
                        }
                    }
                }
            }
            else
            {
                string? val = arg.Value != null ? arg.Value.ToString() : string.Empty;
                if (ignoreEmpty && string.IsNullOrEmpty(val))
                {
                    continue;
                }

                str.Append(val);
                str.Append(delimiter);
                if (str.Length > this.MaxReturnLength)
                {
                    return this.CreateResult(eErrorType.Value);
                }
            }
        }
        string? resultString = str.ToString().TrimEnd(delimiter.ToCharArray());
        return this.CreateResult(resultString, DataType.String);
    }
}