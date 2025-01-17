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

using OfficeOpenXml.Constants;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.Utils;
using System;
using System.Xml;
using static OfficeOpenXml.ExcelWorksheet;

namespace OfficeOpenXml.DataValidation;

/// <summary>
/// A validation containing a formula
/// </summary>
/// <typeparam name="T"></typeparam>
public abstract class ExcelDataValidationWithFormula<T> : ExcelDataValidation
    where T : IExcelDataValidationFormula
{
    protected string _workSheetName;

    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="workSheetName"></param>
    /// <param name="uid">Uid of the data validation, format should be a Guid surrounded by curly braces.</param>
    /// <param name="address"></param>
    internal ExcelDataValidationWithFormula(string uid, string address, ExcelWorksheet ws)
        : base(uid, address, ws) =>
        this._workSheetName = ws.Name;

    /// <summary>
    /// Constructor for reading data
    /// </summary>
    /// <param name="xr">The XmlReader to read from</param>
    internal ExcelDataValidationWithFormula(XmlReader xr, ExcelWorksheet ws)
        : base(xr, ws)
    {
    }

    /// <summary>
    /// Copy Constructor
    /// </summary>
    /// <param name="copy"></param>
    internal ExcelDataValidationWithFormula(ExcelDataValidation copy, ExcelWorksheet ws)
        : base(copy, ws)
    {
    }

    internal override void ReadClassSpecificXmlNodes(XmlReader xr)
    {
        base.ReadClassSpecificXmlNodes(xr);
        this.Formula = this.ReadFormula(xr, "formula1");
    }

    internal T ReadFormula(XmlReader xr, string formulaIdentifier)
    {
        _ = xr.ReadUntil(formulaIdentifier, "dataValidation", "extLst");

        if (xr.LocalName != formulaIdentifier && formulaIdentifier != "formula2")
        {
            throw new NullReferenceException($"Cannot find DataValidation formula for {this.Uid}. " + $"Missing node name: {formulaIdentifier}");
        }

        bool isExt = xr.NamespaceURI == ExcelPackage.schemaMainX14;

        if (this.InternalValidationType == InternalValidationType.ExtLst || isExt)
        {
            _ = xr.Read();
        }

        return this.DefineFormulaClassType(xr.ReadString(), this._workSheetName);
    }

    abstract internal T DefineFormulaClassType(string formulaValue, string worksheetName);

    /// <summary>
    /// Formula - Either a {T} value (except for custom validation) or a spreadsheet formula
    /// </summary>
    public T Formula { get; protected set; }

    /// <summary>
    /// Validates the configuration of the validation.
    /// </summary>
    /// <exception cref="InvalidOperationException">
    /// Will be thrown if invalid configuration of the validation. Details will be in the message of the exception.
    /// </exception>
    public override void Validate()
    {
        base.Validate();

        ExcelDataValidationFormula? formula = this.Formula as ExcelDataValidationFormula;

        if (this.ValidationType.Type == eDataValidationType.Any)
        {
            return;
        }

        if (this.ValidationType.Type == eDataValidationType.List)
        {
            ExcelDataValidationFormulaList? formulaList = (ExcelDataValidationFormulaList)formula;

            if (string.IsNullOrEmpty(formulaList.ExcelFormula) && formulaList.Values.Count <= 0)
            {
                throw new
                    InvalidOperationException($"Formula in Datavalidation of type: {eDataValidationType.List} with Uid: {this.Uid} must have a Value or ExcelFormula");
            }

            return;
        }

        if (formula.HasValue == false && string.IsNullOrEmpty(formula.ExcelFormula) && !(this.AllowBlank ?? false))
        {
            throw new
                InvalidOperationException($"Formula in Datavalidation of type: {this.ValidationType.Type} with Uid: {this.Uid} must have a Value or ExcelFormula");
        }
    }
}