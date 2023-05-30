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

internal class ExcelDataValidationFormulaDateTime : ExcelDataValidationFormulaValue<DateTime?>, IExcelDataValidationFormulaDateTime
{
    public ExcelDataValidationFormulaDateTime(string formula, string validationUid, string sheetName, Action<OnFormulaChangedEventArgs> evtHandler)
        : base(validationUid, sheetName, evtHandler)
    {
        if (!string.IsNullOrEmpty(formula))
        {
            if (double.TryParse(formula, NumberStyles.Any, CultureInfo.InvariantCulture, out double oADate))
            {
                this.Value = DateTime.FromOADate(oADate);
            }
            else
            {
                this.ExcelFormula = formula;
            }
        }
    }

    protected override string GetValueAsString() => this.Value.HasValue ? this.Value.Value.ToOADate().ToString(CultureInfo.InvariantCulture) : string.Empty;
}