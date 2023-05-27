/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;

[FunctionMetadata(Category = ExcelFunctionCategory.Statistical,
                  EPPlusVersion = "5.5",
                  Description = "Returns the average of the absolute deviations of data points from their mean")]
internal class Avedev : ExcelFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        IEnumerable<ExcelDoubleCellValue>? arr = this.ArgsToDoubleEnumerable(arguments, context);

        if (!arr.Any())
        {
            return this.CreateResult(eErrorType.Div0);
        }

        IEnumerable<double>? dArr = arr.Select(x => (double)x);
        double mean = dArr.Average();
        double result = dArr.Aggregate(0d, (val, x) => val += System.Math.Abs(x - mean)) / dArr.Count();

        return this.CreateResult(result, DataType.Decimal);
    }
}