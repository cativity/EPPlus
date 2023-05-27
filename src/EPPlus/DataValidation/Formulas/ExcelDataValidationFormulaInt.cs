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

internal class ExcelDataValidationFormulaInt : ExcelDataValidationFormulaValue<int?>, IExcelDataValidationFormulaInt
{
    public ExcelDataValidationFormulaInt(string formula, string validationUid, string worksheetName, Action<OnFormulaChangedEventArgs> extListHandler)
        : base(validationUid, worksheetName, extListHandler)
    {
        if (!string.IsNullOrEmpty(formula))
        {
            if (int.TryParse(formula, NumberStyles.Number, CultureInfo.InvariantCulture, out int intValue))
            {
                this.Value = intValue;
            }
            else
            {
                this.ExcelFormula = formula;
            }
        }
    }

    protected override string GetValueAsString()
    {
        return this.Value.HasValue ? this.Value.Value.ToString() : string.Empty;
    }
}