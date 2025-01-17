﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/04/2022         EPPlus Software AB       Initial release EPPlus 6.1
 *************************************************************************************************/

using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Export.ToCollection.Exceptions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace OfficeOpenXml.Export.ToCollection;

/// <summary>
/// An object that represents a row in the callback function in <see cref="ExcelRangeBase.ToCollectionWithMappings{T}(Func{ToCollectionRow, T}, ToCollectionRangeOptions)"/>
/// </summary>
public class ToCollectionRow
{
    private ExcelWorkbook _workbook;
    private ToCollectionConversionFailureStrategy _failureStrategy;

    internal ToCollectionRow(List<string> headers, ExcelWorkbook workbook, ToCollectionConversionFailureStrategy failureStrategy)
    {
        this._workbook = workbook;
        this._failureStrategy = failureStrategy;
        this.Headers = headers;

        for (int i = 0; i < headers.Count; i++)
        {
            this._headers.Add(headers[i], i);
        }
    }

    internal Dictionary<string, int> _headers = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);

    /// <summary>
    /// Headers used to access cell values.
    /// </summary>
    public List<string> Headers { get; }

    /// <summary>
    /// The rows values
    /// </summary>
    internal List<ExcelValue> _cellValues { get; set; }

    /// <summary>
    /// Returns the value of the row at the column index
    /// </summary>
    /// <param name="index">the column index</param>
    /// <returns></returns>
    public object this[int index]
    {
        get
        {
            if (index < 0 || index >= this._cellValues.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            return this._cellValues[index];
        }
    }

    /// <summary>
    /// Returns the value of the row at the column index
    /// </summary>
    /// <param name="columnName">the column index</param>
    /// <returns></returns>
    public object this[string columnName]
    {
        get
        {
            if (this._headers.Count == 0)
            {
                throw new ArgumentException($"This range has no headers. Please specify headers or use GetValue<T>(int).");
            }

            if (this._headers.ContainsKey(columnName) == false)
            {
                throw new ArgumentException($"Column name {columnName} does not exist in the range.");
            }

            return this[this._headers[columnName]];
        }
    }

    /// <summary>
    /// Returns the typed value of the cell at the column index within the row of the range.
    /// </summary>
    /// <typeparam name="T">The type to convert to</typeparam>
    /// <param name="index">The column index</param>
    /// <returns>The value</returns>
    /// <exception cref="EPPlusDataTypeConvertionException">Returned if the data type conversion fails and <see cref="ToCollectionOptions.ConversionFailureStrategy"/> is set to Exception</exception>
    public T GetValue<T>(int index)
    {
        if (index < 0 || index >= this._cellValues.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        try
        {
            return ConvertUtil.GetTypedCellValue<T>(this._cellValues[index]._value);
        }
        catch (Exception ex)
        {
            if (this._failureStrategy == ToCollectionConversionFailureStrategy.Exception)
            {
                throw new EPPlusDataTypeConvertionException($"Failure to convert value {this._cellValues[index]._value} for index {index}", ex);
            }

            return default;
        }
    }

    /// <summary>
    /// Returns the typed value of the cell at the column index within the row of the range.
    /// </summary>
    /// <typeparam name="T">The type to convert to</typeparam>
    /// <param name="columnName">The column name</param>
    /// <returns>The value</returns>
    /// <exception cref="EPPlusDataTypeConvertionException">Returned if the data type conversion fails and <see cref="ToCollectionOptions.ConversionFailureStrategy"/> is set to Exception</exception>
    public T GetValue<T>(string columnName)
    {
        if (this._headers.Count == 0)
        {
            throw new ArgumentException($"This range has no headers. Please specify headers or use GetValue<T>(int).");
        }

        if (this._headers.ContainsKey(columnName) == false)
        {
            throw new ArgumentException($"Column name {columnName} does not exist in the range.");
        }

        return this.GetValue<T>(this._headers[columnName]);
    }

    /// <summary>
    /// Returns formatted value of the cell at the column index within the row of the range.
    /// </summary>
    /// <param name="index">The column index</param>
    /// <returns>The formatted value</returns>
    public string GetText(int index)
    {
        if (index < 0 || index >= this._cellValues.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        return ValueToTextHandler.GetFormattedText(this._cellValues[index]._value, this._workbook, this._cellValues[index]._styleId, false);
    }

    /// <summary>
    /// Returns formatted value of the cell at the column index within the row of the range.
    /// </summary>
    /// <param name="columnName">The column name</param>
    /// <returns>The formatted value</returns>
    public string GetText(string columnName)
    {
        if (this._headers.Count == 0)
        {
            throw new ArgumentException($"This range has no headers. Please specify headers or use GetValue<T>(int).");
        }

        if (this._headers.ContainsKey(columnName) == false)
        {
            throw new ArgumentException($"Column name {columnName} does not exist in the range.");
        }

        return this.GetText(this._headers[columnName]);
    }

    List<MappedProperty> _members;

    /// <summary>
    /// Maps properties on the item to values matching the column header with the property name or attibutes without white spaces.
    /// The attributes that can be used are: EpplusTableColumnAttributeBase.Header, DescriptionAttribute.Description or DisplayNameAttribute.Name.
    /// </summary>
    /// <typeparam name="T">The type used, must be a class</typeparam>
    /// <param name="item">The item to set the values on.</param>
    public void Automap<T>(T item)
        where T : class
    {
        this._members ??= ToCollectionAutomap.GetAutomapList<T>(this.Headers);

        foreach (MappedProperty? m in this._members)
        {
            if (m.Index < this._cellValues.Count)
            {
                try
                {
                    m.PropertyInfo.SetValue(item, this._cellValues[m.Index]._value, null);
                }
                catch (Exception ex)
                {
                    if (this._failureStrategy == ToCollectionConversionFailureStrategy.Exception)
                    {
                        EPPlusDataTypeConvertionException? dtcExeption =
                            new
                                EPPlusDataTypeConvertionException($"Can not convert item {this._cellValues[m.Index]._value} to datatype {m.PropertyInfo.DeclaringType}",
                                                                  ex);

                        throw dtcExeption;
                    }
                    else
                    {
                        //Set the default value
                        if (m.PropertyInfo.DeclaringType.IsValueType)
                        {
                            m.PropertyInfo.SetValue(item, null, null);
                        }
                        else
                        {
                            m.PropertyInfo.SetValue(item, Activator.CreateInstance(m.PropertyInfo.DeclaringType), null);
                        }
                    }
                }
            }
        }
    }
}