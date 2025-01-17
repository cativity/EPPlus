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

using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System;
using System.Xml;

namespace OfficeOpenXml.DataValidation;

/// <summary>
/// Represents a data validation with two formulas
/// </summary>
/// <typeparam name="T">An instance implementing the <see cref="IExcelDataValidationFormula"></see></typeparam>
public abstract class ExcelDataValidationWithFormula2<T> : ExcelDataValidationWithFormula<T>
    where T : IExcelDataValidationFormula
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="workSheetName"></param>
    /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
    /// <param name="address"></param>
    internal ExcelDataValidationWithFormula2(string uid, string address, ExcelWorksheet ws)
        : base(uid, address, ws)
    {
    }

    /// <summary>
    /// Constructor for reading data
    /// </summary>
    /// <param name="xr">The XmlReader to read from</param>
    internal ExcelDataValidationWithFormula2(XmlReader xr, ExcelWorksheet ws)
        : base(xr, ws)
    {
    }

    /// <summary>
    /// Copy Constructor
    /// </summary>
    /// <param name="copy"></param>
    internal ExcelDataValidationWithFormula2(ExcelDataValidation copy, ExcelWorksheet ws)
        : base(copy, ws)
    {
    }

    internal override void ReadClassSpecificXmlNodes(XmlReader xr)
    {
        base.ReadClassSpecificXmlNodes(xr);

        if (this.Operator == ExcelDataValidationOperator.between || this.Operator == ExcelDataValidationOperator.notBetween)
        {
            this.Formula2 = this.ReadFormula(xr, "formula2");
        }
        else
        {
            //Ensure Formula2 is not null for writing and if operator type changes.
            this.Formula2 = this.DefineFormulaClassType("", this._workSheetName);
        }
    }

    /// <summary>
    /// Formula - Either a {T} value or a spreadsheet formula
    /// </summary>
    public T Formula2 { get; protected set; }

    //internal override string GetXmlValue()
    //{
    //    base.get
    //}

    public override void Validate()
    {
        base.Validate();

        if (this.ValidationType.Type != eDataValidationType.List
            && this.ValidationType.Type != eDataValidationType.Custom
            && (this.Operator == ExcelDataValidationOperator.between || this.Operator == ExcelDataValidationOperator.notBetween))
        {
            if (string.IsNullOrEmpty(this.Formula2.ExcelFormula)
                && (this.Formula2 as ExcelDataValidationFormula).HasValue == false
                && !(this.AllowBlank ?? false))
            {
                throw new InvalidOperationException("Validation of "
                                                    + this.Address.Address
                                                    + " failed: Formula2 must be set if operator is 'between' or 'notBetween'");
            }
        }
    }
}