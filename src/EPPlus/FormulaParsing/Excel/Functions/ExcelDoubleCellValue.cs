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
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions;

[DebuggerDisplay("{" + nameof(Value) + "}")]
public struct ExcelDoubleCellValue : IComparable<ExcelDoubleCellValue>, IComparable
{
    public ExcelDoubleCellValue(double val)
    {
        this.Value = val;
        this.CellRow = default(int?);
        this.CellCol = default(int?);
    }

    public ExcelDoubleCellValue(double val, int cellRow, int cellCol)
    {
        this.Value = val;
        this.CellRow = cellRow;
        this.CellCol = cellCol;
    }

    public int? CellRow;

    public int? CellCol;

    public double Value;

    public static implicit operator double(ExcelDoubleCellValue d) => d.Value;

    //  User-defined conversion from double to Digit
    public static implicit operator ExcelDoubleCellValue(double d) => new(d);

    public int CompareTo(ExcelDoubleCellValue other) => this.Value.CompareTo(other.Value);

    public int CompareTo(object obj)
    {
        if (obj is double)
        {
            return this.Value.CompareTo((double)obj);
        }

        return this.Value.CompareTo(((ExcelDoubleCellValue)obj).Value);
    }

    public override bool Equals(object obj) => this.CompareTo(obj) == 0;

    public override int GetHashCode() => base.GetHashCode();

    public static bool operator ==(ExcelDoubleCellValue a, ExcelDoubleCellValue b) => a.Value.CompareTo(b.Value) == 0d;

    public static bool operator ==(ExcelDoubleCellValue a, double b) => a.Value.CompareTo(b) == 0d;

    public static bool operator !=(ExcelDoubleCellValue a, ExcelDoubleCellValue b) => a.Value.CompareTo(b.Value) != 0d;

    public static bool operator !=(ExcelDoubleCellValue a, double b) => a.Value.CompareTo(b) != 0d;
}