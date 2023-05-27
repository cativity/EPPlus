/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/15/2020         EPPlus Software AB       ToDataTable function
 *************************************************************************************************/

using OfficeOpenXml.FormulaParsing.Utilities;
using ConvertUtility = OfficeOpenXml.Utils.ConvertUtil;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Reflection;

namespace OfficeOpenXml.Export.ToDataTable;

internal class DataTableExporter
{
    public DataTableExporter(ToDataTableOptions options, ExcelRangeBase range, DataTable dataTable)
    {
        Require.That(options).IsNotNull();
        Require.That(range).IsNotNull();
        Require.That(dataTable).IsNotNull();
        this._options = options;
        this._range = range;
        this._sheet = this._range.Worksheet;
        this._dataTable = dataTable;
    }

    private readonly ToDataTableOptions _options;
    private readonly ExcelRangeBase _range;
    private readonly ExcelWorksheet _sheet;
    private readonly DataTable _dataTable;
    private Dictionary<Type, MethodInfo> _convertMethods = new Dictionary<Type, MethodInfo>();

    public void Export()
    {
        int row = this._options.FirstRowIsColumnNames ? this._range.Start.Row + 1 : this._range.Start.Row;
        this.Validate();
        row += this._options.SkipNumberOfRowsStart;

        while (row <= this._range.End.Row - this._options.SkipNumberOfRowsEnd)
        {
            DataRow? dataRow = this._dataTable.NewRow();
            bool ignoreRow = false;
            bool rowIsEmpty = true;
            string? rowErrorMsg = string.Empty;
            bool rowErrorExists = false;

            foreach (DataColumnMapping? mapping in this._options.Mappings)
            {
                int col = mapping.ZeroBasedColumnIndexInRange + this._range.Start.Column;
                object? val = this._sheet.GetValueInner(row, col);

                if (val != null && rowIsEmpty)
                {
                    rowIsEmpty = false;
                }

                if (!mapping.AllowNull && val == null)
                {
                    rowErrorMsg = $"Value cannot be null, row: {row}, col: {col}";
                    rowErrorExists = true;
                }
                else if (ExcelErrorValue.Values.IsErrorValue(val))
                {
                    switch (this._options.ExcelErrorParsingStrategy)
                    {
                        case ExcelErrorParsingStrategy.IgnoreRowWithErrors:
                            ignoreRow = true;

                            continue;

                        case ExcelErrorParsingStrategy.ThrowException:
                            throw new InvalidOperationException($"Excel error value {val.ToString()} detected at row: {row}, col: {col}");

                        default:
                            val = null;

                            break;
                    }
                }

                if (mapping.TransformCellValue != null)
                {
                    val = mapping.TransformCellValue.Invoke(val);
                }

                Type? type = mapping.ColumnDataType ?? this._dataTable.Columns[mapping.DataColumnName].DataType;
                dataRow[mapping.DataColumnName] = this.CastToColumnDataType(val, type);
            }

            if (rowIsEmpty)
            {
                if (this._options.EmptyRowStrategy == EmptyRowsStrategy.StopAtFirst)
                {
                    row++;

                    break;
                }
            }
            else
            {
                if (rowErrorExists)
                {
                    throw new InvalidOperationException(rowErrorMsg);
                }

                if (!ignoreRow)
                {
                    this._dataTable.Rows.Add(dataRow);
                }
            }

            row++;
        }
    }

    private void Validate()
    {
        int startRow = this._options.FirstRowIsColumnNames ? this._range.Start.Row + 1 : this._range.Start.Row;

        if (this._options.SkipNumberOfRowsStart < 0 || this._options.SkipNumberOfRowsStart > this._range.End.Row - startRow)
        {
            throw new IndexOutOfRangeException("SkipNumberOfRowsStart was out of range: " + this._options.SkipNumberOfRowsStart);
        }

        if (this._options.SkipNumberOfRowsEnd < 0 || this._options.SkipNumberOfRowsEnd > this._range.End.Row - startRow)
        {
            throw new IndexOutOfRangeException("SkipNumberOfRowsEnd was out of range: " + this._options.SkipNumberOfRowsEnd);
        }

        if (this._options.SkipNumberOfRowsEnd + this._options.SkipNumberOfRowsStart > this._range.End.Row - startRow)
        {
            throw new ArgumentException("Total number of skipped rows was larger than number of rows in range");
        }
    }

    private object CastToColumnDataType(object val, Type dataColumnType)
    {
        if (val == null)
        {
            if (dataColumnType.IsValueType)
            {
                return Activator.CreateInstance(dataColumnType);
            }

            return null;
        }

        if (val.GetType() == dataColumnType)
        {
            return val;
        }
        else if (dataColumnType == typeof(DateTime))
        {
            return ConvertUtility.GetValueDate(val);
        }
        else if (dataColumnType == typeof(double))
        {
            return ConvertUtility.GetValueDouble(val);
        }
        else
        {
            try
            {
                if (!this._convertMethods.ContainsKey(dataColumnType))
                {
                    MethodInfo methodInfo = typeof(ConvertUtility).GetMethod(nameof(ConvertUtility.GetTypedCellValue));
                    this._convertMethods.Add(dataColumnType, methodInfo.MakeGenericMethod(dataColumnType));
                }

                MethodInfo? getTypedCellValue = this._convertMethods[dataColumnType];

                return getTypedCellValue.Invoke(null, new object[] { val });
            }
            catch
            {
                if (dataColumnType.IsValueType)
                {
                    return Activator.CreateInstance(dataColumnType);
                }

                return null;
            }
        }
    }
}