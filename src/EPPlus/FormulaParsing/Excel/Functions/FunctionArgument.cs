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
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

public class FunctionArgument
{
    public FunctionArgument(object val)
    {
        this.Value = val;
        this.DataType = DataType.Unknown;
    }

    public FunctionArgument(object val, DataType dataType)
        :this(val)
    {
        this.DataType = dataType;
    }

    private ExcelCellState _excelCellState;

    public void SetExcelStateFlag(ExcelCellState state)
    {
        this._excelCellState |= state;
    }

    public bool ExcelStateFlagIsSet(ExcelCellState state)
    {
        return (this._excelCellState & state) != 0;
    }

    public object Value { get; private set; }

    public DataType DataType { get; }

    public Type Type
    {
        get { return this.Value != null ? this.Value.GetType() : null; }
    }

    public int ExcelAddressReferenceId { get; set; }

    public bool IsExcelRange
    {
        get { return this.Value != null && this.Value is IRangeInfo; }
    }

    public bool IsEnumerableOfFuncArgs
    {
        get { return this.Value != null && this.Value is IEnumerable<FunctionArgument>; }
    }

    public IEnumerable<FunctionArgument> ValueAsEnumerableOfFuncArgs
    {
        get { return this.Value as IEnumerable<FunctionArgument>; }
    }

    public bool ValueIsExcelError
    {
        get { return ExcelErrorValue.Values.IsErrorValue(this.Value); }
    }

    public ExcelErrorValue ValueAsExcelErrorValue
    {
        get { return ExcelErrorValue.Parse(this.Value.ToString()); }
    }

    public IRangeInfo ValueAsRangeInfo
    {
        get { return this.Value as IRangeInfo; }
    }
    public object ValueFirst
    {
        get
        {
            if (this.Value is INameInfo)
            {
                this.Value = ((INameInfo)this.Value).Value;
            }
            IRangeInfo? v = this.Value as IRangeInfo;
            if (v==null)
            {
                return this.Value;
            }
            else
            {
                return v.GetValue(v.Address._fromRow, v.Address._fromCol);
            }
        }
    }

    public string ValueFirstString
    {
        get
        {
            object? v = this.ValueFirst;
            if (v == null)
            {
                return default(string);
            }

            return this.ValueFirst.ToString();
        }
    }

}