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
using OfficeOpenXml.LoadFunctions.Params;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.Export.ToDataTable;

internal class DataTableBuilder
{
    public DataTableBuilder(ToDataTableOptions options, ExcelRangeBase range)
        : this(options, range, null) { }
    public DataTableBuilder(ToDataTableOptions options, ExcelRangeBase range, DataTable dataTable)
    {
        Require.That(options).IsNotNull();
        Require.That(range).IsNotNull();
        this._options = options;
        this._range = range;
        this._sheet = this._range.Worksheet;
        this._dataTable = dataTable;
    }

    private readonly ToDataTableOptions _options;
    private readonly ExcelRangeBase _range;
    private readonly ExcelWorksheet _sheet;
    private DataTable _dataTable;

    internal DataTable Build()
    {
        HashSet<string>? columnNames = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);
        this._dataTable ??= string.IsNullOrEmpty(this._options.DataTableName) ? new DataTable() : new DataTable(this._options.DataTableName);
        if(!string.IsNullOrEmpty(this._options.DataTableNamespace))
        {
            this._dataTable.Namespace = this._options.DataTableNamespace;
        }
        int columnOrder = 0;
        for (int col = this._range.Start.Column; col <= this._range.End.Column; col++)
        {
            int row = this._range.Start.Row;
            string? name = this._options.ColumnNamePrefix + ++columnOrder;
            string? origName = name;
            int columnIndex = columnOrder - 1;
            if(this._options.Mappings.ContainsMapping(columnIndex))
            {
                name = this._options.Mappings.GetByRangeIndex(columnIndex).DataColumnName;
            }
            else if (this._options.FirstRowIsColumnNames)
            {                    
                name = this._sheet.GetValue(row, col)?.ToString();
                origName = name;
                if (name == null)
                {
                    throw new InvalidOperationException(string.Format("First row contains an empty cell at index {0}", col - this._range.Start.Column));
                }

                name = this.GetColumnName(name);
            }
            else
            {
                row--;
            }
            if(columnNames.Contains(name))
            {
                throw new InvalidOperationException($"Duplicate column name : {name}");
            }
            columnNames.Add(name);
            // find type
            while (this._sheet.GetValue(++row, col) == null && row <= this._range.End.Row)
            {
                ;
            }

            object? v = this._sheet.GetValue(row, col);
            if (row == this._range.End.Row && v == null)
            {
                throw new InvalidOperationException(string.Format("Column with index {0} does not contain any values", col));
            }

            Type? type = v == null ? typeof(Nullable) : v.GetType();

            // check mapping
            DataColumnMapping? mapping = this._options.Mappings.GetByRangeIndex(columnIndex);
            if (this._options.PredefinedMappingsOnly && mapping == null)
            {
                continue;
            }
            else if (mapping != null)
            {
                if(mapping.ColumnDataType != null)
                {
                    type = mapping.ColumnDataType;
                }
                if(mapping.HasDataColumn && this._dataTable.Columns[mapping.DataColumnName] == null)
                {
                    this._dataTable.Columns.Add(mapping.DataColumn);
                }
            }

            if((mapping == null || !mapping.HasDataColumn) && this._dataTable.Columns[name] == null)
            {
                DataColumn? column = this._dataTable.Columns.Add(name, type);
                column.Caption = origName;
            }

            if (!this._options.Mappings.ContainsMapping(columnIndex))
            {
                bool allowNull = !type.IsValueType || Nullable.GetUnderlyingType(type) != null;
                this._options.Mappings.Add(columnOrder - 1, name, type, allowNull);
            }
            else if(this._options.Mappings.GetByRangeIndex(columnIndex).ColumnDataType == null)
            {
                this._options.Mappings.GetByRangeIndex(columnIndex).ColumnDataType = type;
            }
        }

        this.HandlePrimaryKeys(this._dataTable);
        return this._dataTable;
    }

    private void HandlePrimaryKeys(DataTable dataTable)
    {
        DataTablePrimaryKey? pk = new DataTablePrimaryKey(this._options);
        if(pk.HasKeys)
        {
            List<DataColumn>? cols = new List<DataColumn>();
            foreach(object? colObj in dataTable.Columns)
            {
                DataColumn? col = colObj as DataColumn;
                if (col == null)
                {
                    continue;
                }

                if (pk.ContainsKey(col.ColumnName))
                {
                    cols.Add(col);
                }   
            }
            dataTable.PrimaryKey = cols.ToArray();
        }
    }

    private string GetColumnName(string name)
    {
        switch(this._options.ColumnNameParsingStrategy)
        {
            case NameParsingStrategy.Preserve:
                return name;
            case NameParsingStrategy.SpaceToUnderscore:
                return name.Replace(" ", "_");
            case NameParsingStrategy.RemoveSpace:
                return name.Replace(" ", string.Empty);
            default:
                return name;
        }
    }
}