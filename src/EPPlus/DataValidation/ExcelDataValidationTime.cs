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

using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System.Xml;

namespace OfficeOpenXml.DataValidation;

/// <summary>
/// Validation for times (<see cref="OfficeOpenXml.DataValidation.ExcelTime"/>).
/// </summary>
public class ExcelDataValidationTime : ExcelDataValidationWithFormula2<IExcelDataValidationFormulaTime>, IExcelDataValidationTime
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="worksheet"></param>
    /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
    /// <param name="address"></param>
    /// <param name="validationType"></param>
    internal ExcelDataValidationTime(string uid, string address, ExcelWorksheet ws)
        : base(uid, address, ws)
    {
        this.Formula = new ExcelDataValidationFormulaTime(null, uid, ws.Name, this.OnFormulaChanged);
        this.Formula2 = new ExcelDataValidationFormulaTime(null, uid, ws.Name, this.OnFormulaChanged);
    }

    /// <summary>
    /// Constructor for reading data
    /// </summary>
    /// <param name="xr">The XmlReader to read from</param>
    internal ExcelDataValidationTime(XmlReader xr, ExcelWorksheet ws)
        : base(xr, ws)
    {
    }

    /// <summary>
    /// Copy constructor
    /// </summary>
    /// <param name="copy"></param>
    internal ExcelDataValidationTime(ExcelDataValidationTime copy, ExcelWorksheet ws)
        : base(copy, ws)
    {
        this.Formula = copy.Formula;
        this.Formula2 = copy.Formula2;
    }

    /// <summary>
    /// Property for determining type of validation
    /// </summary>
    public override ExcelDataValidationType ValidationType => new ExcelDataValidationType(eDataValidationType.Time);

    internal override IExcelDataValidationFormulaTime DefineFormulaClassType(string formulaValue, string sheetName) => new ExcelDataValidationFormulaTime(formulaValue, this.Uid, sheetName, this.OnFormulaChanged);

    internal override ExcelDataValidation GetClone() => new ExcelDataValidationTime(this, this._ws);

    internal override ExcelDataValidation GetClone(ExcelWorksheet copy) => new ExcelDataValidationTime(this, copy);

    ExcelDataValidationTime Clone() => (ExcelDataValidationTime)this.GetClone();
}