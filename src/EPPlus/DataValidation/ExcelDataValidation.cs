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
using OfficeOpenXml.DataValidation.Events;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Xml;

namespace OfficeOpenXml.DataValidation;

/// <summary>
/// Abstract base class for all Excel datavalidations. Contains functionlity which is common for all these different validation types.
/// </summary>
public abstract class ExcelDataValidation : IExcelDataValidation
{
    /// <summary>
    /// Constructor
    /// </summary>
    /// <param name="uid">Id for validation</param>
    /// <param name="address">adress validation is applied to</param>
    protected ExcelDataValidation(string uid, string address, ExcelWorksheet ws)
    {
        Require.Argument(uid).IsNotNullOrEmpty("uid");
        Require.Argument(address).IsNotNullOrEmpty("address");

        this.Uid = uid;
        this._address = new ExcelDatavalidationAddress(CheckAndFixRangeAddress(address), this);
        this._ws = ws;
    }

    /// <summary>
    /// Read-File Constructor
    /// </summary>
    /// <param name="xr"></param>
    protected ExcelDataValidation(XmlReader xr, ExcelWorksheet ws)
    {
        this.LoadXML(xr);
        this._ws = ws;
    }

    /// <summary>
    /// Copy-Constructor
    /// </summary>
    /// <param name="validation">Validation to copy from</param>
    protected ExcelDataValidation(ExcelDataValidation validation, ExcelWorksheet ws)
    {
        this.Uid = validation.Uid;
        this.Address = validation.Address;
        this.ValidationType = validation.ValidationType;
        this.ErrorStyle = validation.ErrorStyle;
        this.AllowBlank = validation.AllowBlank;
        this.ShowInputMessage = validation.ShowInputMessage;
        this.ShowErrorMessage = validation.ShowErrorMessage;
        this.ErrorTitle = validation.ErrorTitle;
        this.Error = validation.Error;
        this.PromptTitle = validation.PromptTitle;
        this.Prompt = validation.Prompt;
        this.operatorString = validation.operatorString;
        this._ws = ws;
    }

    internal ExcelWorksheet _ws;

    /// <summary>
    /// Uid of the data validation
    /// </summary>
    public string Uid { get; internal set; }

    ExcelDatavalidationAddress _address;

    /// <summary>
    /// Address of data validation
    /// </summary>
    public ExcelAddress Address
    {
        get => this._address;
        internal set => this._address = (ExcelDatavalidationAddress)value;
    }

    /// <summary>
    /// Validation type
    /// </summary>
    public virtual ExcelDataValidationType ValidationType { get; }

    string errorStyleString;

    /// <summary>
    /// Warning style
    /// </summary>
    public ExcelDataValidationWarningStyle ErrorStyle
    {
        get
        {
            if (!string.IsNullOrEmpty(this.errorStyleString))
            {
                return (ExcelDataValidationWarningStyle)Enum.Parse(typeof(ExcelDataValidationWarningStyle), this.errorStyleString, true);
            }

            return ExcelDataValidationWarningStyle.undefined;
        }
        set
        {
            if (value == ExcelDataValidationWarningStyle.undefined)
            {
                this.errorStyleString = null;
            }
            else
            {
                this.errorStyleString = value.ToString();
            }
        }
    }

    string imeModeString;

    public ExcelDataValidationImeMode ImeMode
    {
        get
        {
            if (string.IsNullOrEmpty(this.imeModeString))
            {
                return ExcelDataValidationImeMode.NoControl;
            }

            return (ExcelDataValidationImeMode)this.imeModeString.ToEnum<ExcelDataValidationImeMode>();
        }
        set
        {
            if (value == ExcelDataValidationImeMode.NoControl)
            {
                this.imeModeString = null;
            }
            else
            {
                this.imeModeString = value.ToString();
            }
        }
    }

    /// <summary>
    /// True if blanks should be allowed
    /// </summary>
    public bool? AllowBlank { get; set; }

    /// <summary>
    /// True if input message should be shown
    /// </summary>
    public bool? ShowInputMessage { get; set; }

    /// <summary>
    /// True if error message should be shown
    /// </summary>
    public bool? ShowErrorMessage { get; set; }

    /// <summary>
    /// Title of error message box
    /// </summary>
    public string ErrorTitle { get; set; }

    /// <summary>
    /// Error message box text
    /// </summary>
    public string Error { get; set; }

    /// <summary>
    /// Title of the validation message box.
    /// </summary>
    public string PromptTitle { get; set; }

    /// <summary>
    /// Text of the validation message box.
    /// </summary>
    public string Prompt { get; set; }

    /// <summary>
    /// True if the current validation type allows operator.
    /// </summary>
    public virtual bool AllowsOperator => true;

    /// <summary>
    /// This method will validate the state of the validation
    /// </summary>
    /// <exception cref="InvalidOperationException">If the state breaks the rules of the validation</exception>
    public virtual void Validate()
    {
    }

    ExcelDataValidationAsType _as;

    /// <summary>
    /// Us this property to case <see cref="IExcelDataValidation"/>s to its subtypes
    /// </summary>
    public ExcelDataValidationAsType As => this._as ??= new ExcelDataValidationAsType(this);

    /// <summary>
    /// Indicates whether this instance is stale, see https://github.com/EPPlusSoftware/EPPlus/wiki/Data-validation-Exceptions
    /// DEPRECATED as of Epplus 6.2.
    /// This as validations can no longer be stale since all attributes are now always fresh and held in the system.
    /// </summary>
    [Obsolete]
    public bool IsStale { get; } = false;

    string operatorString;

    /// <summary>
    /// Operator for comparison between the entered value and Formula/Formulas.
    /// </summary>
    public ExcelDataValidationOperator Operator
    {
        get
        {
            if (!string.IsNullOrEmpty(this.operatorString))
            {
                return (ExcelDataValidationOperator)Enum.Parse(typeof(ExcelDataValidationOperator), this.operatorString, true);
            }

            return default(ExcelDataValidationOperator);
        }
        set
        {
            if (this.ValidationType.Type == eDataValidationType.Any || this.ValidationType.Type == eDataValidationType.List)
            {
                throw new InvalidOperationException("The current validation type does not allow operator to be set");
            }

            this.operatorString = value.ToString();
        }
    }

    private static string CheckAndFixRangeAddress(string address)
    {
        if (address.Contains(","))
        {
            throw new FormatException("Multiple addresses may not be commaseparated, use space instead");
        }

        address = ConvertUtil._invariantTextInfo.ToUpper(address);

        if (IsEntireColumn(address))
        {
            address = AddressUtility.ParseEntireColumnSelections(address);
        }

        return address;
    }

    static bool IsEntireColumn(string address)
    {
        bool hasColon = false;

        foreach (char c in address)
        {
            if ((c >= 'A' && c <= 'Z') || c == ':')
            {
                if (c == ':')
                {
                    hasColon = true;
                }

                continue;
            }
            else
            {
                return false;
            }
        }

        if (hasColon)
        {
            return true;
        }
        else
        {
            return false;
        }
    }

    /// <summary>
    /// Type to determine if extLst or not
    /// </summary>
    internal InternalValidationType InternalValidationType { get; set; } = InternalValidationType.DataValidation;

    /// <summary>
    /// Event method for changing internal type when referring to an external worksheet.
    /// </summary>
    protected Action<OnFormulaChangedEventArgs> OnFormulaChanged =>
        (e) =>
        {
            if (e.isExt)
            {
                this.InternalValidationType = InternalValidationType.ExtLst;
            }
        };

    internal virtual void LoadXML(XmlReader xr)
    {
        string address = xr.GetAttribute("sqref");

        if (address == null)
        {
            this.InternalValidationType = InternalValidationType.ExtLst;
        }

        this.Uid = string.IsNullOrEmpty(xr.GetAttribute("xr:uid")) ? NewId() : xr.GetAttribute("xr:uid");

        this.operatorString = xr.GetAttribute("operator");
        this.errorStyleString = xr.GetAttribute("errorStyle");

        this.imeModeString = xr.GetAttribute("imeMode");

        this.AllowBlank = xr.GetAttribute("allowBlank") == "1" ? true : false;

        this.ShowInputMessage = xr.GetAttribute("showInputMessage") == "1" ? true : false;
        this.ShowErrorMessage = xr.GetAttribute("showErrorMessage") == "1" ? true : false;

        this.ErrorTitle = xr.GetAttribute("errorTitle");
        this.Error = xr.GetAttribute("error");

        this.PromptTitle = xr.GetAttribute("promptTitle");
        this.Prompt = xr.GetAttribute("prompt");

        this.ReadClassSpecificXmlNodes(xr);

        if (address == null && xr.ReadUntil(5, "sqref", "dataValidation", "extLst"))
        {
            address = xr.ReadString();

            if (address == null)
            {
                throw new NullReferenceException($"Unable to locate ExtList adress for DataValidation with uid:{this.Uid}");
            }
        }

        this._address = new ExcelDatavalidationAddress(CheckAndFixRangeAddress(address).Replace(" ", ","), this);
    }

    internal virtual void ReadClassSpecificXmlNodes(XmlReader xr)
    {
    }

    internal static string NewId() => "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";

    internal void SetAddress(string address)
    {
        _ = AddressUtility.ParseEntireColumnSelections(address);
        this._address = new ExcelDatavalidationAddress(address, this);
        this._ws.DataValidations.UpdateRangeDictionary(this);
    }

    /// <summary>
    /// Create a Deep-Copy of this validation.
    /// Note that one should also implement a separate clone() method casting to the child class
    /// </summary>
    internal abstract ExcelDataValidation GetClone();

    /// <summary>
    /// Create a Deep-Copy of this validation.
    /// Note that one should also implement a separate clone() method casting to the child class
    /// </summary>
    internal abstract ExcelDataValidation GetClone(ExcelWorksheet copy);
}