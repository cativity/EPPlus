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

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

[FunctionMetadata(Category = ExcelFunctionCategory.Text, EPPlusVersion = "4", Description = "Converts a text string into a numeric value")]
internal class Value : ExcelFunction
{
    public Value(CultureInfo ci)
    {
        this._cultureInfo = ci;
        this._groupSeparator = this._cultureInfo.NumberFormat.NumberGroupSeparator;
        this._decimalSeparator = this._cultureInfo.NumberFormat.NumberDecimalSeparator;
        this._timeSeparator = this._cultureInfo.DateTimeFormat.TimeSeparator;
        //this._shortTimePattern = this._cultureInfo.DateTimeFormat.ShortTimePattern;
    }

    private readonly CultureInfo _cultureInfo;
    private readonly string _groupSeparator;
    private readonly string _decimalSeparator;
    private readonly string _timeSeparator;
    //private readonly string _shortTimePattern;
    private readonly DateValue _dateValueFunc = new DateValue();
    private readonly TimeValue _timeValueFunc = new TimeValue();

    public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
    {
        ValidateArguments(arguments, 1);
        string? val = ArgToString(arguments, 0);
        double result = 0d;

        if (string.IsNullOrEmpty(val))
        {
            return this.CreateResult(result, DataType.Integer);
        }

        val = val.TrimEnd(' ');
        bool isPercentage = false;

        if (val.EndsWith("%"))
        {
            val = val.TrimEnd('%');
            isPercentage = true;
        }

        if (val.StartsWith("(", StringComparison.OrdinalIgnoreCase) && val.EndsWith(")", StringComparison.OrdinalIgnoreCase))
        {
            string? numCandidate = val.Substring(1, val.Length - 2);

            if (double.TryParse(numCandidate, NumberStyles.Any, this._cultureInfo, out double _))
            {
                val = "-" + numCandidate;
            }
        }

        if (Regex.IsMatch(val,
                          $"^[\\d]*({Regex.Escape(this._groupSeparator)}?[\\d]*)*?({Regex.Escape(this._decimalSeparator)}[\\d]*)?[ ?% ?]?$",
                          RegexOptions.Compiled))
        {
            result = double.Parse(val, this._cultureInfo);

            return this.CreateResult(isPercentage ? result / 100 : result, DataType.Decimal);
        }

        if (double.TryParse(val, NumberStyles.Float, this._cultureInfo, out result))
        {
            return this.CreateResult(isPercentage ? result / 100d : result, DataType.Decimal);
        }

        string? timeSeparator = Regex.Escape(this._timeSeparator);

        if (Regex.IsMatch(val, @"^[\d]{1,2}" + timeSeparator + @"[\d]{2}(" + timeSeparator + @"[\d]{2})?$", RegexOptions.Compiled))
        {
            CompileResult? timeResult = this._timeValueFunc.Execute(val);

            if (timeResult.DataType == DataType.Date)
            {
                return timeResult;
            }
        }

        CompileResult? dateResult = this._dateValueFunc.Execute(val);

        if (dateResult.DataType == DataType.Date)
        {
            return dateResult;
        }

        return this.CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
    }
}