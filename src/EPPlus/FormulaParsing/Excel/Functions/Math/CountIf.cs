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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Statistical,
        EPPlusVersion = "4",
        Description = "Returns the number of cells (of a supplied range), that satisfy a given criteria")]
    internal class CountIf : ExcelFunction
    {
        private readonly ExpressionEvaluator _expressionEvaluator;

        public CountIf()
            : this(new ExpressionEvaluator())
        {

        }

        public CountIf(ExpressionEvaluator evaluator)
        {
            Require.That(evaluator).Named("evaluator").IsNotNull();
            this._expressionEvaluator = evaluator;
        }

        private bool Evaluate(object obj, string expression)
        {
            double? candidate = default(double?);
            if (IsNumeric(obj))
            {
                candidate = ConvertUtil.GetValueDouble(obj);
            }
            if (candidate.HasValue)
            {
                return this._expressionEvaluator.Evaluate(candidate.Value, expression, false);
            }
            return this._expressionEvaluator.Evaluate(obj, expression, false);
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            FunctionArgument[]? functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            ValidateArguments(functionArguments, 2);
            FunctionArgument? range = functionArguments.ElementAt(0);
            string? criteria = functionArguments.ElementAt(1).ValueFirstString;
            double result = 0d;
            if (range.IsExcelRange)
            {
                IRangeInfo? rangeInfo = range.ValueAsRangeInfo;
                for (int row = rangeInfo.Address.Start.Row; row < rangeInfo.Address.End.Row + 1; row++)
                {
                    for (int col = rangeInfo.Address.Start.Column; col < rangeInfo.Address.End.Column + 1; col++)
                    {
                        if (criteria != null && this.Evaluate(rangeInfo.GetValue(row, col), criteria))
                        {
                            result++;
                        }
                    }
                }
            }
            else if (range.Value is IEnumerable<FunctionArgument>)
            {
                foreach (FunctionArgument? arg in (IEnumerable<FunctionArgument>) range.Value)
                {
                    if(this.Evaluate(arg.Value, criteria))
                    {
                        result++;
                    }
                }
            }
            else
            {
                if (this.Evaluate(range.Value, criteria))
                {
                    result++;
                }
            }
            return this.CreateResult(result, DataType.Integer);
        }
    }
}
