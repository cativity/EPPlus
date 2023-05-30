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
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable;

/// <summary>
/// Class used to map columns in the <see cref="ExcelRangeBase.ToDataTable(ToDataTableOptions, DataTable)"/> method
/// </summary>
public class DataColumnMapping
{
    internal DataColumnMapping(DataColumn dataColumn)
    {
        Require.That(dataColumn).IsNotNull();
        Require.That(dataColumn.ColumnName).IsNotNullOrEmpty();
        Require.That(dataColumn.DataType).IsNotNull();
        this.DataColumn = dataColumn;
    }

    internal DataColumnMapping()
    {
    }

    internal bool HasDataColumn => this.DataColumn != null;

    /// <summary>
    /// The <see cref="System.Data.DataColumn"/> used for the mapping
    /// </summary>
    public DataColumn DataColumn { get; private set; }

    /// <summary>
    /// Zero based index of the mappings column in the range
    /// </summary>
    public int ZeroBasedColumnIndexInRange { get; set; }

    private string _dataColumnName;

    /// <summary>
    /// Name of the data column, corresponds to <see cref="System.Data.DataColumn.ColumnName"/>
    /// </summary>
    public string DataColumnName
    {
        get => this.HasDataColumn ? this.DataColumn.ColumnName : this._dataColumnName;
        set
        {
            if (this.HasDataColumn)
            {
                this.DataColumn.ColumnName = value;
            }
            else
            {
                this._dataColumnName = value;
            }
        }
    }

    private Type _dataColumnType;

    /// <summary>
    /// <see cref="Type">Type</see> of the column, corresponds to <see cref="System.Data.DataColumn.DataType"/>
    /// </summary>
    public Type ColumnDataType
    {
        get
        {
            if (this.HasDataColumn)
            {
                return this.DataColumn.DataType;
            }
            else
            {
                return this._dataColumnType;
            }
        }
        set
        {
            if (this.HasDataColumn)
            {
                this.DataColumn.DataType = value;
            }
            else
            {
                this._dataColumnType = value;
            }
        }
    }

    private bool _allowNull;

    /// <summary>
    /// Indicates whether empty cell values should be allowed. Corresponds to <see cref="System.Data.DataColumn.AllowDBNull"/>
    /// </summary>
    public bool AllowNull
    {
        get
        {
            if (this.HasDataColumn)
            {
                return this.DataColumn.AllowDBNull;
            }
            else
            {
                return this._allowNull;
            }
        }
        set
        {
            if (this.HasDataColumn)
            {
                this.DataColumn.AllowDBNull = value;
            }
            else
            {
                this._allowNull = value;
            }
        }
    }

    /// <summary>
    /// A function which allows casting of an <see cref="Object"/> before it is written to the <see cref="DataTable"/>
    /// </summary>
    /// <example>
    /// <code>
    /// var options = ToDataTableOptions.Create(o =>
    /// {
    ///     // the last argument is a lambda function that will call the read value's ToString method
    ///     // and this string will be written to the DataTable
    ///     o.Mappings.Add(0, "Id", typeof(string), true, c => "Id: " + c.ToString());
    /// });
    /// </code>
    /// </example>
    public Func<object, object> TransformCellValue { get; set; }

    internal void Validate()
    {
        if (string.IsNullOrEmpty(this.DataColumnName))
        {
            throw new ArgumentNullException("DataColumnName");
        }

        if (this.ZeroBasedColumnIndexInRange < 0)
        {
            throw new ArgumentOutOfRangeException("ZeroBasedColumnIndex");
        }
    }
}