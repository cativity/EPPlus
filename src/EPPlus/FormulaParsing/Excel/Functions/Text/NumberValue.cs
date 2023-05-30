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

using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

[FunctionMetadata(Category = ExcelFunctionCategory.Text,
                  EPPlusVersion = "5.0",
                  Description = "Converts text to a number, in a locale-independent way",
                  IntroducedInExcelVersion = "2013")]
internal class NumberValue : ExcelFunction
{
    private string _decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
    private string _groupSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
    private string _arg = string.Empty;
    private int _nPercentage;

    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        string? arg = ArgToString(arguments, 0);
        this.SetArgAndPercentage(arg);

        if (!this.ValidateAndSetSeparators(arguments.ToArray()))
        {
            return this.CreateResult(ExcelErrorValue.Values.Value, DataType.ExcelError);
        }

        CultureInfo? cultureInfo = new CultureInfo("en-US", true);
        cultureInfo.NumberFormat.NumberDecimalSeparator = this._decimalSeparator;
        cultureInfo.NumberFormat.NumberGroupSeparator = this._groupSeparator;

        if (double.TryParse(this._arg, NumberStyles.Any, cultureInfo, out double result))
        {
            if (this._nPercentage > 0)
            {
                result /= System.Math.Pow(100, this._nPercentage);
            }

            return this.CreateResult(result, DataType.Decimal);
        }

        return this.CreateResult(ExcelErrorValue.Values.Value, DataType.ExcelError);
    }

    private void SetArgAndPercentage(string arg)
    {
        int pIndex = arg.IndexOf("%", StringComparison.OrdinalIgnoreCase);

        if (pIndex > 0)
        {
            this._arg = arg.Substring(0, pIndex).Replace(" ", "");
            string? percentage = arg.Substring(pIndex, arg.Length - pIndex).Trim();

            if (!Regex.IsMatch(percentage, "[%]+"))
            {
                throw new ArgumentException("Invalid format: " + arg);
            }

            this._nPercentage = percentage.Length;
        }
        else
        {
            this._arg = arg;
        }
    }

    private bool ValidateAndSetSeparators(FunctionArgument[] arguments)
    {
        if (arguments.Length == 1)
        {
            return true;
        }

        string? decimalSeparator = ArgToString(arguments, 1).Substring(0, 1);

        if (!DecimalSeparatorIsValid(decimalSeparator))
        {
            return false;
        }

        this._decimalSeparator = decimalSeparator;

        if (arguments.Length > 2)
        {
            string? groupSeparator = ArgToString(arguments, 2).Substring(0, 1);

            if (!GroupSeparatorIsValid(decimalSeparator, groupSeparator))
            {
                return false;
            }

            this._groupSeparator = groupSeparator;
        }

        return true;
    }

    private static bool DecimalSeparatorIsValid(string separator)
    {
        return !string.IsNullOrEmpty(separator) && (separator == "." || separator == ",");
    }

    private static bool GroupSeparatorIsValid(string groupSeparator, string decimalSeparator)
    {
        return !string.IsNullOrEmpty(groupSeparator)
               && groupSeparator != decimalSeparator
               && (groupSeparator == " " || groupSeparator == "," || groupSeparator == ".");
    }
}