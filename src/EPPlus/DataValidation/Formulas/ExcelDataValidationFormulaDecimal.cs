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

using OfficeOpenXml.DataValidation.Events;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System;
using System.Globalization;

namespace OfficeOpenXml.DataValidation.Formulas;

/// <summary>
/// 
/// </summary>
internal class ExcelDataValidationFormulaDecimal : ExcelDataValidationFormulaValue<double?>, IExcelDataValidationFormulaDecimal
{
    public ExcelDataValidationFormulaDecimal(string formula, string validationUid, string sheetName, Action<OnFormulaChangedEventArgs> extHandler)
        : base(validationUid, sheetName, extHandler)
    {
        if (!string.IsNullOrEmpty(formula))
        {
            if (double.TryParse(formula, NumberStyles.Any, CultureInfo.InvariantCulture, out double dValue))
            {
                this.Value = dValue;
            }
            else
            {
                this.ExcelFormula = formula;
            }
        }
    }

    protected override string GetValueAsString()
    {
        return this.Value.HasValue ? this.Value.Value.ToString("R15", CultureInfo.InvariantCulture) : string.Empty;
    }
}