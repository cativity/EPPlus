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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.DataValidation;

internal class RangeDataValidation : IRangeDataValidation
{
    public RangeDataValidation(ExcelWorksheet worksheet, string address)
    {
        Require.Argument(worksheet).IsNotNull("worksheet");
        Require.Argument(address).IsNotNullOrEmpty("address");
        this._worksheet = worksheet;
        this._address = address;
    }

    ExcelWorksheet _worksheet;
    string _address;

    /// <summary>
    ///  Used to remove all dataValidations in cell or cellrange
    /// </summary>
    /// <param name="deleteIfEmpty">Deletes the dataValidation if it has no addresses after clear</param>
    /// <exception cref="InvalidOperationException"></exception>
    public void ClearDataValidation(bool deleteIfEmpty = false)
    {
        ExcelAddress? address = new ExcelAddress(this._address);

        List<ExcelDataValidation>? validations =
            this._worksheet.DataValidations._validationsRD.GetValuesFromRange(address._fromRow, address._fromCol, address._toRow, address._toCol);

        foreach (ExcelDataValidation? validation in validations)
        {
            ExcelAddressBase? excelAddress = new ExcelAddressBase(validation.Address.Address.Replace(" ", ","));
            List<ExcelAddressBase>? addresses = excelAddress.GetAllAddresses();

            string newAddress = "";

            foreach (ExcelAddressBase? validationAddress in addresses)
            {
                ExcelAddressBase? nullOrAddress = validationAddress.IntersectReversed(address);

                if (nullOrAddress != null)
                {
                    newAddress += nullOrAddress.Address + " ";
                }
            }

            if (newAddress == "")
            {
                if (deleteIfEmpty)
                {
                    _ = this._worksheet.DataValidations.Remove(validation);
                }
                else
                {
                    throw new InvalidOperationException($"Cannot remove last address in validation of type {validation.ValidationType.Type} "
                                                        + $"with uid {validation.Uid} without deleting it."
                                                        + $" Add other addresses or use ClearDataValidation(true)");
                }
            }
            else
            {
                validation.Address.Address = newAddress;
            }
        }
    }

    public IExcelDataValidationAny AddAnyDataValidation()
    {
        return this._worksheet.DataValidations.AddAnyValidation(this._address);
    }

    public IExcelDataValidationInt AddIntegerDataValidation()
    {
        return this._worksheet.DataValidations.AddIntegerValidation(this._address);
    }

    public IExcelDataValidationDecimal AddDecimalDataValidation()
    {
        return this._worksheet.DataValidations.AddDecimalValidation(this._address);
    }

    public IExcelDataValidationDateTime AddDateTimeDataValidation()
    {
        return this._worksheet.DataValidations.AddDateTimeValidation(this._address);
    }

    public IExcelDataValidationList AddListDataValidation()
    {
        return this._worksheet.DataValidations.AddListValidation(this._address);
    }

    public IExcelDataValidationInt AddTextLengthDataValidation()
    {
        return this._worksheet.DataValidations.AddTextLengthValidation(this._address);
    }

    public IExcelDataValidationTime AddTimeDataValidation()
    {
        return this._worksheet.DataValidations.AddTimeValidation(this._address);
    }

    public IExcelDataValidationCustom AddCustomDataValidation()
    {
        return this._worksheet.DataValidations.AddCustomValidation(this._address);
    }
}