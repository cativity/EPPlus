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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

internal abstract class MultipleRangeCriteriasFunction : ExcelFunction
{
    private readonly ExpressionEvaluator _expressionEvaluator;

    protected MultipleRangeCriteriasFunction()
        : this(new ExpressionEvaluator())
    {
    }

    protected MultipleRangeCriteriasFunction(ExpressionEvaluator evaluator)
    {
        Require.That(evaluator).Named("evaluator").IsNotNull();
        this._expressionEvaluator = evaluator;
    }

    protected bool Evaluate(object obj, string expression, bool convertNumericString = true)
    {
        double? candidate = default(double?);

        if (IsNumeric(obj))
        {
            candidate = ConvertUtil.GetValueDouble(obj);
        }

        if (candidate.HasValue)
        {
            return this._expressionEvaluator.Evaluate(candidate.Value, expression, convertNumericString);
        }

        return this._expressionEvaluator.Evaluate(obj, expression, convertNumericString);
    }

    protected List<int> GetMatchIndexes(RangeOrValue rangeOrValue, string searched, bool convertNumericString = true)
    {
        List<int>? result = new List<int>();
        int internalIndex = 0;

        if (rangeOrValue.Range != null)
        {
            IRangeInfo? rangeInfo = rangeOrValue.Range;
            int toRow = rangeInfo.Address._toRow;

            if (rangeInfo.Worksheet.Dimension.End.Row < toRow)
            {
                toRow = rangeInfo.Worksheet.Dimension.End.Row;
            }

            for (int row = rangeInfo.Address._fromRow; row <= toRow; row++)
            {
                for (int col = rangeInfo.Address._fromCol; col <= rangeInfo.Address._toCol; col++)
                {
                    object? candidate = rangeInfo.GetValue(row, col);

                    if (searched != null && this.Evaluate(candidate, searched, convertNumericString))
                    {
                        result.Add(internalIndex);
                    }

                    internalIndex++;
                }
            }
        }
        else if (this.Evaluate(rangeOrValue.Value, searched, convertNumericString))
        {
            result.Add(internalIndex);
        }

        return result;
    }
}