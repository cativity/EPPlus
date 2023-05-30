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

namespace OfficeOpenXml.DataValidation.Formulas;

/// <summary>
/// 
/// </summary>
internal class ExcelDataValidationFormulaCustom : ExcelDataValidationFormula, IExcelDataValidationFormula
{
    public ExcelDataValidationFormulaCustom(string formula, string validationUid, string sheetName, Action<OnFormulaChangedEventArgs> extHandler)
        : base(validationUid, sheetName, extHandler)
    {
        if (!string.IsNullOrEmpty(formula))
        {
            this.ExcelFormula = formula;
        }

        this.State = FormulaState.Formula;
    }

    internal override string GetXmlValue() => this.ExcelFormula;

    protected override string GetValueAsString() => this.ExcelFormula;

    internal override void ResetValue()
    {
    }
}