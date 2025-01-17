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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

internal class RowMatcher
{
    private readonly WildCardValueMatcher _wildCardValueMatcher;
    private readonly ExpressionEvaluator _expressionEvaluator;

    public RowMatcher()
        : this(new WildCardValueMatcher(), new ExpressionEvaluator())
    {
    }

    public RowMatcher(WildCardValueMatcher wildCardValueMatcher, ExpressionEvaluator expressionEvaluator)
    {
        this._wildCardValueMatcher = wildCardValueMatcher;
        this._expressionEvaluator = expressionEvaluator;
    }

    public bool IsMatch(ExcelDatabaseRow row, ExcelDatabaseCriteria criteria)
    {
        bool retVal = true;

        foreach (KeyValuePair<ExcelDatabaseCriteriaField, object> c in criteria.Items)
        {
            object? candidate = c.Key.FieldIndex.HasValue ? row[c.Key.FieldIndex.Value] : row[c.Key.FieldName];
            object? crit = c.Value;

            if (candidate.IsNumeric() && crit.IsNumeric())
            {
                if (System.Math.Abs(ConvertUtil.GetValueDouble(candidate) - ConvertUtil.GetValueDouble(crit)) > double.Epsilon)
                {
                    return false;
                }
            }
            else
            {
                string? criteriaString = crit.ToString();

                if (!this.Evaluate(candidate, criteriaString))
                {
                    return false;
                }
            }
        }

        return retVal;
    }

    private bool Evaluate(object obj, string expression)
    {
        if (obj == null)
        {
            return false;
        }

        double? candidate = default(double?);

        if (ConvertUtil.IsNumericOrDate(obj))
        {
            candidate = ConvertUtil.GetValueDouble(obj);
        }

        if (candidate.HasValue)
        {
            return this._expressionEvaluator.Evaluate(candidate.Value, expression);
        }

        return this._wildCardValueMatcher.IsMatch(expression, obj.ToString()) == 0;
    }
}