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

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

[FunctionMetadata(Category = ExcelFunctionCategory.Statistical, EPPlusVersion = "4", Description = "Returns the largest value from a list of supplied numbers")]
internal class Median : HiddenValuesHandlingFunction
{
    public Median()
    {
        this.IgnoreErrors = false;
    }

    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        IEnumerable<ExcelDoubleCellValue>? nums = this.ArgsToDoubleEnumerable(this.IgnoreHiddenValues, this.IgnoreErrors, arguments, context);
        ExcelDoubleCellValue[]? arr = nums.ToArray();
        Array.Sort(arr);

        if (arr.Length == 0)
        {
            return this.CreateResult(eErrorType.Num);
        }

        double result;

        if (arr.Length % 2 == 1)
        {
            result = arr[arr.Length / 2];
        }
        else
        {
            int startIndex = (arr.Length / 2) - 1;
            result = (arr[startIndex] + arr[startIndex + 1]) / 2d;
        }

        return this.CreateResult(result, DataType.Decimal);
    }
}