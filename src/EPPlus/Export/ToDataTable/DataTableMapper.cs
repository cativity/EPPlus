using OfficeOpenXml.FormulaParsing.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable
{
    internal class DataTableMapper
    {
        public DataTableMapper(ToDataTableOptions options, ExcelRangeBase range, DataTable dataTable)
        {
            Require.That(options).IsNotNull();
            Require.That(dataTable).IsNotNull();
            Require.That(range).IsNotNull();
            this._options = options;
            this._dataTable = dataTable;
            this._range = range;
        }

        private readonly ToDataTableOptions _options;
        private readonly DataTable _dataTable;
        private readonly ExcelRangeBase _range;

        internal void Map()
        {
            int indexInRange = 0;
            foreach(object? columnObj in this._dataTable.Columns)
            {
                DataColumn? column = columnObj as DataColumn;
                if (column == null)
                {
                    continue;
                }

                if(!this._options.Mappings.Any(x => string.Compare(column.ColumnName, x.DataColumnName, StringComparison.OrdinalIgnoreCase) == 0))
                {
                    if(this._options.FirstRowIsColumnNames)
                    {
                        int ix = this.FindIndexInRange(column.ColumnName);
                        if (ix == -1)
                        {
                            throw new InvalidOperationException("Column name not found in range: " + column.ColumnName);
                        }

                        this._options.Mappings.Add(ix, column.ColumnName, column.DataType, column.AllowDBNull);
                    }
                    else
                    {
                        this._options.Mappings.Add(indexInRange, column.ColumnName, column.DataType, column.AllowDBNull);
                    }
                    indexInRange++;
                }
            }
        }

        private int FindIndexInRange(string columnName)
        {
            int row = this._range.Start.Row;
            int index = 0;
            for(int col = this._range.Start.Column; col <= this._range.End.Column; col++)
            {
                object? cellVal = this._range.Worksheet.GetValueInner(row, col);
                if (cellVal == null)
                {
                    continue;
                }

                if (string.Compare(columnName, cellVal.ToString(), StringComparison.OrdinalIgnoreCase) == 0)
                {
                    return index;
                }
                index++;
            }
            return -1;
        }
    }
}
