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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

[FunctionMetadata(Category = ExcelFunctionCategory.Statistical,
                  EPPlusVersion = "4",
                  Description = "Calculates the Average of the cells in a supplied range, that satisfy a given criteria",
                  IntroducedInExcelVersion = "2007")]
internal class AverageIf : HiddenValuesHandlingFunction
{
    private readonly ExpressionEvaluator _expressionEvaluator;

    public AverageIf()
        : this(new ExpressionEvaluator())
    {
    }

    public AverageIf(ExpressionEvaluator evaluator)
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
            return this._expressionEvaluator.Evaluate(candidate.Value, expression);
        }

        return this._expressionEvaluator.Evaluate(obj, expression);
    }

    private static string GetCriteraFromArg(IEnumerable<FunctionArgument> arguments) => arguments.ElementAt(1).ValueFirst != null ? ArgToString(arguments, 1) : null;

    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 2);
        IRangeInfo? argRange = ArgToRangeInfo(arguments, 0);
        string? criteria = GetCriteraFromArg(arguments);
        double returnValue;

        if (argRange == null)
        {
            object? val = arguments.ElementAt(0).Value;

            if (criteria != null && this.Evaluate(val, criteria))
            {
                IRangeInfo? lookupRange = ArgToRangeInfo(arguments, 2);
                returnValue = arguments.Count() > 2 ? lookupRange.First().ValueDouble : ConvertUtil.GetValueDouble(val, true);
            }
            else
            {
                throw new ExcelErrorValueException(eErrorType.Div0);
            }
        }
        else if (arguments.Count() > 2)
        {
            IRangeInfo? lookupRange = ArgToRangeInfo(arguments, 2);
            returnValue = this.CalculateWithLookupRange(argRange, criteria, lookupRange, context);
        }
        else
        {
            returnValue = this.CalculateSingleRange(argRange, criteria, context);
        }

        return this.CreateResult(returnValue, DataType.Decimal);
    }

    private double CalculateWithLookupRange(IRangeInfo argRange, string criteria, IRangeInfo sumRange, ParsingContext context)
    {
        double returnValue = 0d;
        int nMatches = 0;

        foreach (ICellInfo? cell in argRange)
        {
            if (criteria != null && this.Evaluate(cell.Value, criteria))
            {
                int rowOffset = cell.Row - argRange.Address._fromRow;
                int columnOffset = cell.Column - argRange.Address._fromCol;

                if (sumRange.Address._fromRow + rowOffset <= sumRange.Address._toRow && sumRange.Address._fromCol + columnOffset <= sumRange.Address._toCol)
                {
                    object? val = sumRange.GetOffset(rowOffset, columnOffset);

                    if (val is ExcelErrorValue)
                    {
                        ThrowExcelErrorValueException((ExcelErrorValue)val);
                    }

                    nMatches++;
                    returnValue += ConvertUtil.GetValueDouble(val, true);
                }
            }
        }

        return Divide(returnValue, nMatches);
    }

    private double CalculateSingleRange(IRangeInfo range, string expression, ParsingContext context)
    {
        double returnValue = 0d;
        int nMatches = 0;

        foreach (ICellInfo? candidate in range)
        {
            if (expression != null && IsNumeric(candidate.Value) && this.Evaluate(candidate.Value, expression))
            {
                if (candidate.IsExcelError)
                {
                    ThrowExcelErrorValueException((ExcelErrorValue)candidate.Value);
                }

                returnValue += candidate.ValueDouble;
                nMatches++;
            }
        }

        return Divide(returnValue, nMatches);
    }
}