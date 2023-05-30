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
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph;

public class CompileResult
{
    private static CompileResult _empty = new CompileResult(null, DataType.Empty);
    private static CompileResult _zeroDecimal = new CompileResult(0d, DataType.Decimal);
    private static CompileResult _zeroInt = new CompileResult(0d, DataType.Integer);

    /// <summary>
    /// Returns a CompileResult with a null value and data type set to DataType.Empty
    /// </summary>
    public static CompileResult Empty => _empty;

    /// <summary>
    /// Returns a CompileResult instance with a decimal value of 0.
    /// </summary>
    public static CompileResult ZeroDecimal => _zeroDecimal;

    /// <summary>
    /// Returns a CompileResult instance with a integer value of 0.
    /// </summary>
    public static CompileResult ZeroInt => _zeroInt;

    private double? _resultNumeric;

    public CompileResult(object result, DataType dataType)
        : this(result, dataType, 0)
    {
    }

    public CompileResult(object result, DataType dataType, int excelAddressReferenceId)
    {
        if (result is ExcelDoubleCellValue)
        {
            this.Result = ((ExcelDoubleCellValue)result).Value;
        }
        else
        {
            this.Result = result;
        }

        this.DataType = dataType;
        this.ExcelAddressReferenceId = excelAddressReferenceId;
    }

    public CompileResult(eErrorType errorType)
    {
        this.Result = ExcelErrorValue.Create(errorType);
        this.DataType = DataType.ExcelError;
    }

    public CompileResult(ExcelErrorValue errorValue)
    {
        Require.Argument(errorValue).IsNotNull("errorValue");
        this.Result = errorValue;
        this.DataType = DataType.ExcelError;
    }

    public object Result { get; private set; }

    public object ResultValue
    {
        get
        {
            IRangeInfo? r = this.Result as IRangeInfo;

            if (r == null)
            {
                return this.Result;
            }
            else
            {
                return r.GetValue(r.Address._fromRow, r.Address._fromCol);
            }
        }
    }

    public double ResultNumeric
    {
        get
        {
            // We assume that Result does not change unless it is a range.
            if (this._resultNumeric == null)
            {
                if (this.IsNumeric)
                {
                    this._resultNumeric = this.Result == null ? 0 : Convert.ToDouble(this.Result);
                }
                else if (this.IsPercentageString && ConvertUtil.TryParsePercentageString(this.Result.ToString(), out double v))
                {
                    this._resultNumeric = v;
                }
                else if (this.Result is DateTime)
                {
                    this._resultNumeric = ((DateTime)this.Result).ToOADate();
                }
                else if (this.Result is TimeSpan)
                {
                    this._resultNumeric = DateTime.FromOADate(0).Add((TimeSpan)this.Result).ToOADate();
                }
                else if (this.Result is IRangeInfo)
                {
                    ICellInfo? c = ((IRangeInfo)this.Result).FirstOrDefault();

                    if (c == null)
                    {
                        return 0;
                    }
                    else
                    {
                        return c.ValueDoubleLogical;
                    }
                }

                // The IsNumericString and IsDateString properties will set _resultNumeric for efficiency so we just need
                // to check them here.
                else if (!this.IsDateString && !this.IsNumericString)
                {
                    this._resultNumeric = 0;
                }
            }

            return this._resultNumeric.Value;
        }
    }

    public DataType DataType { get; private set; }

    public bool IsNumeric =>
        this.DataType == DataType.Decimal
        || this.DataType == DataType.Integer
        || this.DataType == DataType.Empty
        || this.DataType == DataType.Boolean
        || this.DataType == DataType.Date
        || this.DataType == DataType.Time;

    public bool IsNumericString
    {
        get
        {
            if (this.DataType == DataType.String && ConvertUtil.TryParseNumericString(this.Result as string, out double result))
            {
                this._resultNumeric = result;

                return true;
            }

            return false;
        }
    }

    public bool IsPercentageString
    {
        get
        {
            if (this.DataType == DataType.String)
            {
                string? s = this.Result as string;

                return ConvertUtil.IsPercentageString(s);
            }

            return false;
        }
    }

    public bool IsDateString
    {
        get
        {
            if (this.DataType == DataType.String && ConvertUtil.TryParseDateString(this.Result as string, out DateTime result))
            {
                this._resultNumeric = result.ToOADate();

                return true;
            }

            return false;
        }
    }

    public bool IsResultOfSubtotal { get; set; }

    public bool IsHiddenCell { get; set; }

    public int ExcelAddressReferenceId { get; set; }

    public bool IsResultOfResolvedExcelRange => this.ExcelAddressReferenceId > 0;
}