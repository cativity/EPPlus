using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;

public abstract class IfsWithMultipleMatchesBase : ExcelFunction
{
    private readonly ExpressionEvaluator _evaluator = new ExpressionEvaluator();
    protected IEnumerable<double> GetMatches(string functionName, IEnumerable<FunctionArgument> arguments, out CompileResult errorResult)
    {
        ValidateArguments(arguments, 3);
        errorResult = null;
        IRangeInfo? maxRange = arguments.ElementAt(0).ValueAsRangeInfo;
        int maxArgs = arguments.Count() < (126 * 2) + 1 ? arguments.Count() : (126 * 2) + 1;
        List<double>? matches = new List<double>();
        bool rangeSizeEvaluated = false;
        for (int valueIx = 0; valueIx < maxRange.Count(); valueIx++)
        {
            bool isMatch = true;
            for (int criteriaIx = 1; criteriaIx < maxArgs; criteriaIx += 2)
            {

                IRangeInfo? criteriaRange = arguments.ElementAt(criteriaIx).ValueAsRangeInfo;
                if (!rangeSizeEvaluated)
                {
                    if (criteriaRange.Count() < maxRange.Count())
                    {
                        errorResult = this.CreateResult(eErrorType.Value);
                        return Enumerable.Empty<double>();
                    }
                }
                object? matchCriteria = arguments.ElementAt(criteriaIx + 1).Value;

                object? candidate = criteriaRange.ElementAt(valueIx).Value;
                if (!this._evaluator.Evaluate(candidate, Convert.ToString(matchCriteria, CultureInfo.InvariantCulture)))
                {
                    isMatch = false;
                    break;
                }
            }
            rangeSizeEvaluated = true;
            if (isMatch)
            {
                matches.Add(maxRange.ElementAt(valueIx).ValueDouble);
            }
        }
        return matches;
    }
}