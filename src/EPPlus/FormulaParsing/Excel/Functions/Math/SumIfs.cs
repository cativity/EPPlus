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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

[FunctionMetadata(Category = ExcelFunctionCategory.MathAndTrig,
                  EPPlusVersion = "4",
                  Description = "Adds the cells in a supplied range, that satisfy multiple criteria",
                  IntroducedInExcelVersion = "2007")]
internal class SumIfs : MultipleRangeCriteriasFunction
{
    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        FunctionArgument[]? functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
        ValidateArguments(functionArguments, 3);
        List<int>? rows = new List<int>();
        IRangeInfo? valueRange = functionArguments[0].ValueAsRangeInfo;
        List<double> sumRange;

        if (valueRange != null)
        {
            sumRange = this.ArgsToDoubleEnumerableZeroPadded(false, valueRange, context).ToList();
        }
        else
        {
            sumRange = this.ArgsToDoubleEnumerable(false, new List<FunctionArgument> { functionArguments[0] }, context).Select(x => (double)x).ToList();
        }

        List<RangeOrValue>? argRanges = new List<RangeOrValue>();
        List<string>? criterias = new List<string>();

        for (int ix = 1; ix < 31; ix += 2)
        {
            if (functionArguments.Length <= ix)
            {
                break;
            }

            FunctionArgument? arg = functionArguments[ix];

            if (arg.IsExcelRange)
            {
                IRangeInfo? rangeInfo = arg.ValueAsRangeInfo;
                argRanges.Add(new RangeOrValue { Range = rangeInfo });
            }
            else
            {
                argRanges.Add(new RangeOrValue { Value = arg.Value });
            }

            string? value = functionArguments[ix + 1].Value != null ? ArgToString(arguments, ix + 1) : null;
            criterias.Add(value);
        }

        IEnumerable<int> matchIndexes = this.GetMatchIndexes(argRanges[0], criterias[0]);
        IList<int>? enumerable = matchIndexes as IList<int> ?? matchIndexes.ToList();

        for (int ix = 1; ix < argRanges.Count && enumerable.Any(); ix++)
        {
            List<int>? indexes = this.GetMatchIndexes(argRanges[ix], criterias[ix]);
            matchIndexes = matchIndexes.Intersect(indexes);
        }

        double result = matchIndexes.Sum(index => sumRange[index]);

        return this.CreateResult(result, DataType.Decimal);
    }
}