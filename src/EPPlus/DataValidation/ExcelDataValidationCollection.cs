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

using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.DataValidation;

/// <summary>
/// <para>
/// Collection of <see cref="ExcelDataValidation"/>. This class is providing the API for EPPlus data validation.
/// </para>
/// <para>
/// The public methods of this class (Add[...]Validation) will create a datavalidation entry in the worksheet. When this
/// validation has been created changes to the properties will affect the workbook immediately.
/// </para>
/// <para>
/// Each type of validation has either a formula or a typed value/values, except for custom validation which has a formula only.
/// </para>
/// <code>
/// // Add a date time validation
/// var validation = worksheet.DataValidation.AddDateTimeValidation("A1");
/// // set validation properties
/// validation.ShowErrorMessage = true;
/// validation.ErrorTitle = "An invalid date was entered";
/// validation.Error = "The date must be between 2011-01-31 and 2011-12-31";
/// validation.Prompt = "Enter date here";
/// validation.Formula.Value = DateTime.Parse("2011-01-01");
/// validation.Formula2.Value = DateTime.Parse("2011-12-31");
/// validation.Operator = ExcelDataValidationOperator.between;
/// </code>
/// </summary>
public class ExcelDataValidationCollection : IEnumerable<IExcelDataValidation>
{
    private List<ExcelDataValidation> _validations = new List<ExcelDataValidation>();
    private ExcelWorksheet _worksheet;
    internal RangeDictionary<ExcelDataValidation> _validationsRD = new RangeDictionary<ExcelDataValidation>();

    internal ExcelDataValidationCollection(ExcelWorksheet worksheet)
    {
        this.InternalValidationEnabled = true;
        this._worksheet = worksheet;
    }

    internal ExcelDataValidationCollection(XmlReader xr, ExcelWorksheet worksheet)
        : this(worksheet)
    {
        this.InternalValidationEnabled = true;
        this.ReadDataValidations(xr);
    }

    /// <summary>
    /// Read data validation from xml via xr reader
    /// </summary>
    public void ReadDataValidations(XmlReader xr)
    {
        while (xr.Read())
        {
            if (xr.LocalName != "dataValidation")
            {
                _ = xr.Read();

                break;
            }

            if (xr.NodeType == XmlNodeType.Element)
            {
                ExcelDataValidation? validation = ExcelDataValidationFactory.Create(xr, this._worksheet);

                if (validation.Address.Addresses != null)
                {
                    for (int i = 0; i < validation.Address.Addresses.Count; i++)
                    {
                        this._validationsRD.Merge(validation.Address.Addresses[i]._fromRow,
                                                  validation.Address.Addresses[i]._fromCol,
                                                  validation.Address.Addresses[i]._toRow,
                                                  validation.Address.Addresses[i]._toCol,
                                                  validation);
                    }
                }
                else
                {
                    this._validationsRD.Merge(validation.Address._fromRow,
                                              validation.Address._fromCol,
                                              validation.Address._toRow,
                                              validation.Address._toCol,
                                              validation);
                }

                this._validations.Add(validation);
            }
        }
    }

    internal void AddToRangeDictionary(ExcelDataValidation validation) => this.AddItemToRangeDictionaryMultipleAddresses(validation.Address.Address, validation);

    internal void UpdateRangeDictionary(ExcelDataValidation validation)
    {
        if (validation.Address.Addresses != null)
        {
            for (int i = 0; i < validation.Address.Addresses.Count; i++)
            {
                this._validationsRD.Merge(validation.Address.Addresses[i]._fromRow,
                                          validation.Address.Addresses[i]._fromCol,
                                          validation.Address.Addresses[i]._toRow,
                                          validation.Address.Addresses[i]._toCol,
                                          validation);
            }
        }
        else
        {
            this._validationsRD.Merge(validation.Address._fromRow,
                                      validation.Address._fromCol,
                                      validation.Address._toRow,
                                      validation.Address._toCol,
                                      validation);
        }
    }

    internal bool HasValidationType(InternalValidationType type)
    {
        if (this.Count != 0)
        {
            for (int i = 0; i < this.Count; i++)
            {
                if (this._validations[i].InternalValidationType == type)
                {
                    return true;
                }
            }

            return false;
        }
        else
        {
            return false;
        }
    }

    int GetCount(InternalValidationType type)
    {
        int validationCount = 0;

        for (int i = 0; i < this.Count; i++)
        {
            if (this._validations[i].InternalValidationType == type)
            {
                validationCount++;
            }
        }

        return validationCount;
    }

    internal int GetNonExtLstCount() => this.GetCount(InternalValidationType.DataValidation);

    internal int GetExtLstCount() => this.GetCount(InternalValidationType.ExtLst);

    private static void OnValidationCountChanged()
    {
    }

    /// <summary>
    /// Validates address - not empty, collisions
    /// </summary>
    /// <param name="address"></param>
    /// <param name="validatingValidation"></param>
    private void ValidateAddress(string address, IExcelDataValidation validatingValidation)
    {
        Require.Argument(address).IsNotNullOrEmpty("address");

        if (!this.InternalValidationEnabled)
        {
            return;
        }

        // ensure that the new address does not collide with an existing validation.
        ExcelAddress? newAddress = new ExcelAddress(address);

        if (this._validations.Count > 0)
        {
            foreach (ExcelDataValidation? validation in this._validations)
            {
                if (validatingValidation != null && validatingValidation == validation)
                {
                    continue;
                }

                ExcelAddressBase.eAddressCollition result = validation.Address.Collide(newAddress);

                if (result != ExcelAddressBase.eAddressCollition.No)
                {
                    throw new InvalidOperationException(string.Format("The address ({0}) collides with an existing validation ({1})",
                                                                      address,
                                                                      validation.Address.Address));
                }
            }
        }
    }

    /// <summary>
    /// Validates all data validations.
    /// </summary>
    internal void ValidateAll()
    {
        if (!this.InternalValidationEnabled)
        {
            return;
        }

        foreach (ExcelDataValidation? validation in this._validations)
        {
            validation.Validate();
        }
    }

    /// <summary>
    /// Optionally add address at end for new copy with address in range
    /// </summary>
    /// <param name="dv"></param>
    /// <param name="address"></param>
    internal void AddCopyOfDataValidation(ExcelDataValidation dv, ExcelWorksheet added, string address = null)
    {
        if (address == null)
        {
            this._validations.Add(dv.GetClone(added));
        }
        else
        {
            this._validations.Add(ExcelDataValidationFactory.CloneWithNewAdress(address, dv, added));
        }
    }

    /// <summary>
    /// Adds a <see cref="ExcelDataValidationAny"/> to the worksheet.
    /// </summary>
    /// <param name="address">The range/address to validate</param>
    /// <returns></returns>
    public IExcelDataValidationAny AddAnyValidation(string address)
    {
        ExcelDataValidationAny? validation = new ExcelDataValidationAny(ExcelDataValidation.NewId(), address, this._worksheet);

        return (IExcelDataValidationAny)this.AddValidation(address, validation);
    }

    /// <summary>
    /// Adds an <see cref="IExcelDataValidationInt"/> to the worksheet. Whole means that the only accepted values
    /// are integer values.
    /// </summary>
    /// <param name="address">the range/address to validate</param>
    public IExcelDataValidationInt AddIntegerValidation(string address)
    {
        ExcelDataValidationInt? validation = new ExcelDataValidationInt(ExcelDataValidation.NewId(), address, this._worksheet);

        return (IExcelDataValidationInt)this.AddValidation(address, validation);
    }

    /// <summary>
    /// Adds an <see cref="IExcelDataValidationInt"/> regarding text length to the worksheet.
    /// </summary>
    /// <param name="address">The range/address to validate</param>
    public IExcelDataValidationInt AddTextLengthValidation(string address)
    {
        ExcelDataValidationInt? validation = new ExcelDataValidationInt(ExcelDataValidation.NewId(), address, this._worksheet, true);

        return (IExcelDataValidationInt)this.AddValidation(address, validation);
    }

    /// <summary>
    /// Addes an <see cref="IExcelDataValidationDecimal"/> to the worksheet. The only accepted values are
    /// decimal values.
    /// </summary>
    /// <param name="address">The range/address to validate</param>
    public IExcelDataValidationDecimal AddDecimalValidation(string address)
    {
        ExcelDataValidationDecimal? validation = new ExcelDataValidationDecimal(ExcelDataValidation.NewId(), address, this._worksheet);

        return (IExcelDataValidationDecimal)this.AddValidation(address, validation);
    }

    /// <summary>
    /// Adds an <see cref="IExcelDataValidationList"/> to the worksheet. The accepted values are defined
    /// in a list.
    /// </summary>
    /// <param name="address">The range/address to validate</param>
    public IExcelDataValidationList AddListValidation(string address)
    {
        ExcelDataValidationList? validation = new ExcelDataValidationList(ExcelDataValidation.NewId(), address, this._worksheet);

        return (IExcelDataValidationList)this.AddValidation(address, validation);
    }

    /// <summary>
    /// Adds an <see cref="IExcelDataValidationDateTime"/> to the worksheet.
    /// </summary>
    /// <param name="address">The range/address to validate</param>
    public IExcelDataValidationDateTime AddDateTimeValidation(string address)
    {
        ExcelDataValidationDateTime? validation = new ExcelDataValidationDateTime(ExcelDataValidation.NewId(), address, this._worksheet);

        return (IExcelDataValidationDateTime)this.AddValidation(address, validation);
    }

    /// <summary>
    /// Adds an <see cref="IExcelDataValidationDateTime"/> to the worksheet.
    /// </summary>
    /// <param name="address">The range/address to validate</param>
    public IExcelDataValidationTime AddTimeValidation(string address)
    {
        ExcelDataValidationTime? validation = new ExcelDataValidationTime(ExcelDataValidation.NewId(), address, this._worksheet);

        return (IExcelDataValidationTime)this.AddValidation(address, validation);
    }

    /// <summary>
    /// Adds a <see cref="ExcelDataValidationCustom"/> to the worksheet.
    /// </summary>
    /// <param name="address">The range/address to validate</param>
    public IExcelDataValidationCustom AddCustomValidation(string address)
    {
        ExcelDataValidationCustom? validation = new ExcelDataValidationCustom(ExcelDataValidation.NewId(), address, this._worksheet);

        return (IExcelDataValidationCustom)this.AddValidation(address, validation);
    }

    private ExcelDataValidation AddValidation(string address, ExcelDataValidation validation)
    {
        this._validations.Add(validation);

        //_ = new ExcelAddress(address.Replace(" ", ","));

        this.AddItemToRangeDictionaryMultipleAddresses(address, validation);

        return validation;
    }

    private void AddItemToRangeDictionaryMultipleAddresses(string address, ExcelDataValidation validation)
    {
        ExcelAddress? internalAddress = new ExcelAddress(address.Replace(" ", ","));

        foreach (ExcelAddressBase? individualAddress in internalAddress.GetAllAddresses())
        {
            if (this._validationsRD.Exists(individualAddress._fromRow, individualAddress._fromCol, individualAddress._toRow, individualAddress._toCol))
            {
                throw new InvalidOperationException($"A DataValidation already exists at {address}"
                                                    + $" If using ClearDataValidation this may be because the sheet you're reading has multiple dataValidations on one cell.");
            }

            this._validationsRD.Add(individualAddress._fromRow, individualAddress._fromCol, individualAddress._toRow, individualAddress._toCol, validation);
        }
    }

    /// <summary>
    /// Number of validations
    /// </summary>3
    public int Count => this._validations.Count;

    /// <summary>
    /// Epplus validates that all data validations are consistend and valid
    /// when they are added and when a workbook is saved. Since this takes some
    /// resources, it can be disabled for improve performance. 
    /// </summary>
    public bool InternalValidationEnabled { get; set; }

    /// <summary>
    /// Index operator, returns by 0-based index
    /// </summary>
    /// <param name="index"></param>
    /// <returns></returns>
    public ExcelDataValidation this[int index]
    {
        get => this._validations[index];
        set => this._validations[index] = value;
    }

    /// <summary>
    /// Index operator, returns a data validation which address partly or exactly matches the searched address.
    /// </summary>
    /// <param name="address">A cell address or range</param>
    /// <returns>A <see cref="ExcelDataValidation"/> or null if no match</returns>
    public IExcelDataValidation this[string address]
    {
        get
        {
            ExcelAddress? searchedAddress = new ExcelAddress(address);

            return this._validations.Find(x => x.Address.Collide(searchedAddress) != ExcelAddressBase.eAddressCollition.No);
        }
    }

    /// <summary>
    /// Returns all validations that matches the supplied predicate <paramref name="match"/>.
    /// </summary>
    /// <param name="match">predicate to filter out matching validations</param>
    /// <returns></returns>
    public IEnumerable<ExcelDataValidation> FindAll(Predicate<ExcelDataValidation> match) => this._validations.FindAll(match);

    /// <summary>
    /// Removes an <see cref="ExcelDataValidation"/> from the collection.
    /// </summary>
    /// <param name="item">The item to remove</param>
    /// <returns>True if remove succeeds, otherwise false</returns>
    /// <exception cref="ArgumentNullException">if <paramref name="item"/> is null</exception>
    public bool Remove(IExcelDataValidation item)
    {
        Require.Argument(item).IsNotNull("item");

        if (!(item is ExcelDataValidation))
        {
            throw new InvalidCastException("The supplied item must inherit OfficeOpenXml.DataValidation.ExcelDataValidation");
        }

        bool retVal = this._validations.Remove((ExcelDataValidation)item);

        if (retVal)
        {
            OnValidationCountChanged();
        }

        return retVal;
    }

    /// <summary>
    /// Returns the first matching validation.
    /// </summary>
    /// <param name="match"></param>
    /// <returns></returns>
    public ExcelDataValidation Find(Predicate<ExcelDataValidation> match) => this._validations.Find(match);

    /// <summary>
    /// Removes all validations from the collection.
    /// </summary>
    public void Clear() => this._validations.Clear();

    /// <summary>
    /// Removes the validations that matches the predicate
    /// </summary>
    /// <param name="match"></param>
    public void RemoveAll(Predicate<ExcelDataValidation> match)
    {
        List<ExcelDataValidation>? matches = this._validations.FindAll(match);

        foreach (ExcelDataValidation? m in matches)
        {
            if (!(m is ExcelDataValidation))
            {
                throw new InvalidCastException("The supplied item must inherit OfficeOpenXml.DataValidation.ExcelDataValidation");
            }
        }

        _ = this._validations.RemoveAll(match);
    }

    IEnumerator<IExcelDataValidation> IEnumerable<IExcelDataValidation>.GetEnumerator()
    {
        for (int i = 0; i < this._validations.Count; i++)
        {
            yield return this._validations[i];
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => this._validations.GetEnumerator();

    internal void InsertRangeDictionary(ExcelAddressBase address, bool shiftRight)
    {
        if (address.Addresses == null)
        {
            this.InsertRangeIntoRangeDictionary(address, shiftRight);
        }
        else
        {
            foreach (ExcelAddressBase? a in address.Addresses)
            {
                this.InsertRangeIntoRangeDictionary(a, shiftRight);
            }
        }
    }

    private void InsertRangeIntoRangeDictionary(ExcelAddressBase address, bool shiftRight)
    {
        if (shiftRight)
        {
            this._validationsRD.InsertColumn(address._fromCol, address.Columns, address._fromRow, address._toRow);
        }
        else
        {
            this._validationsRD.InsertRow(address._fromRow, address.Rows, address._fromCol, address._toCol);
        }
    }

    internal void ClearRangeDictionary(ExcelAddressBase address)
    {
        ExcelAddressBase? internalAddress = new ExcelAddressBase(address.Address.Replace(" ", ","));

        foreach (ExcelAddressBase? individualAddress in internalAddress.GetAllAddresses())
        {
            this._validationsRD.DeleteRow(individualAddress._fromRow, individualAddress.Rows, individualAddress._fromCol, individualAddress._toCol, false);
        }
    }

    internal void DeleteRangeDictionary(ExcelAddressBase address, bool shiftLeft)
    {
        if (address.Addresses == null)
        {
            this.DeleteRangeInRangeDictionary(address, shiftLeft);
        }
        else
        {
            foreach (ExcelAddressBase? a in address.Addresses)
            {
                this.DeleteRangeInRangeDictionary(a, shiftLeft);
            }
        }
    }

    private void DeleteRangeInRangeDictionary(ExcelAddressBase address, bool shiftLeft)
    {
        if (shiftLeft)
        {
            this._validationsRD.DeleteColumn(address._fromCol, address.Columns, address._fromRow, address._toRow);
        }
        else
        {
            this._validationsRD.DeleteRow(address._fromRow, address.Rows, address._fromCol, address._toCol);
        }
    }
}