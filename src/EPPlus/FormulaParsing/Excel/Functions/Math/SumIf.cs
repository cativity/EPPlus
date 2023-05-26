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
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.MathAndTrig,
        EPPlusVersion = "4",
        Description = "Adds the cells in a supplied range, that satisfy a given criteria")]
    internal class SumIf : HiddenValuesHandlingFunction
    {
        private readonly ExpressionEvaluator _evaluator;

        public SumIf()
            : this(new ExpressionEvaluator())
        {

        }

        public SumIf(ExpressionEvaluator evaluator)
        {
            Require.That(evaluator).Named("evaluator").IsNotNull();
            this._evaluator = evaluator;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            IRangeInfo? argRange = ArgToRangeInfo(arguments, 0);

            // Criteria can either be a string or an array of strings
            IEnumerable<string>? criteria = GetCriteria(arguments.ElementAt(1));
            double retVal = 0d;
            if (argRange == null)
            {
                object? val = arguments.ElementAt(0).Value;
                if (this._evaluator.Evaluate(val, criteria))
                {
                    if (arguments.Count() > 2)
                    {
                        IRangeInfo? sumRange = ArgToRangeInfo(arguments, 2);
                        retVal = sumRange.First().ValueDouble;
                    }
                    else
                    {
                        retVal = ConvertUtil.GetValueDouble(val, true);
                    }
                }
            }
            else if (arguments.Count() > 2)
            {
                IRangeInfo? sumRange = ArgToRangeInfo(arguments, 2);
                retVal = this.CalculateWithSumRange(argRange, criteria, sumRange, context);
            }
            else
            {
                retVal = this.CalculateSingleRange(argRange, criteria, context);
            }
            return this.CreateResult(retVal, DataType.Decimal);
        }

        internal static IEnumerable<string> GetCriteria(FunctionArgument criteriaArg)
        {
            List<string>? criteria = new List<string>();
            if (criteriaArg.IsEnumerableOfFuncArgs)
            {
                foreach (FunctionArgument? arg in criteriaArg.ValueAsEnumerableOfFuncArgs)
                {
                    criteria.Add(arg.ValueFirstString);
                }
            }
            else if (criteriaArg.IsExcelRange)
            {
                foreach (ICellInfo? cell in criteriaArg.ValueAsRangeInfo)
                {
                    if (cell.Value != null)
                    {
                        criteria.Add(cell.Value.ToString());
                    }
                }
            }
            else
            {
                criteria.Add(criteriaArg.ValueFirst != null ? criteriaArg.ValueFirst.ToString() : null);
            }
            return criteria;
        }

        private double CalculateWithSumRange(IRangeInfo range, IEnumerable<string> criteria, IRangeInfo sumRange, ParsingContext context)
        {
            double retVal = 0d;
            foreach (ICellInfo? cell in range)
            {
                if (this._evaluator.Evaluate(cell.Value, criteria))
                {
                    int rowOffset = cell.Row - range.Address._fromRow;
                    int columnOffset = cell.Column - range.Address._fromCol;
                    if (sumRange.Address._fromRow + rowOffset <= sumRange.Address._toRow &&
                       sumRange.Address._fromCol + columnOffset <= sumRange.Address._toCol)
                    {
                        object? val = sumRange.GetOffset(rowOffset, columnOffset);
                        if (val is ExcelErrorValue)
                        {
                            ThrowExcelErrorValueException((ExcelErrorValue)val);
                        }
                        retVal += ConvertUtil.GetValueDouble(val, true);
                    }
                }
            }
            return retVal;
        }

        private double CalculateSingleRange(IRangeInfo range, IEnumerable<string> expressions, ParsingContext context)
        {
            double retVal = 0d;
            foreach (ICellInfo? candidate in range)
            {
                if (IsNumeric(candidate.Value) && this._evaluator.Evaluate(candidate.Value, expressions) && IsNumeric(candidate.Value))
                {
                    if (candidate.IsExcelError)
                    {
                        ThrowExcelErrorValueException((ExcelErrorValue)candidate.Value);
                    }
                    retVal += candidate.ValueDouble;
                }
            }
            return retVal;
        }
    }
}
