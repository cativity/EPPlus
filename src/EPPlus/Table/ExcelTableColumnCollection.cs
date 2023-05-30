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

using OfficeOpenXml.Core;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table;

/// <summary>
/// A collection of table columns
/// </summary>
public class ExcelTableColumnCollection : IEnumerable<ExcelTableColumn>
{
    List<ExcelTableColumn> _cols = new List<ExcelTableColumn>();
    Dictionary<string, int> _colNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
    internal int _maxId = 1;

    internal ExcelTableColumnCollection(ExcelTable table)
    {
        this.Table = table;

        foreach (XmlNode node in table.TableXml.SelectNodes("//d:table/d:tableColumns/d:tableColumn", table.NameSpaceManager))
        {
            ExcelTableColumn? item = new ExcelTableColumn(table.NameSpaceManager, node, table, this._cols.Count);
            this._cols.Add(item);
            this._colNames.Add(this._cols[this._cols.Count - 1].Name, this._cols.Count - 1);
            int id = item.Id;

            if (id >= this._maxId)
            {
                this._maxId = id + 1;
            }
        }
    }

    /// <summary>
    /// A reference to the table object
    /// </summary>
    public ExcelTable Table { get; private set; }

    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count => this._cols.Count;

    /// <summary>
    /// The column Index. Base 0.
    /// </summary>
    /// <param name="Index"></param>
    /// <returns></returns>
    public ExcelTableColumn this[int Index]
    {
        get
        {
            if (Index < 0 || Index >= this._cols.Count)
            {
                throw new ArgumentOutOfRangeException("Column index out of range");
            }

            return this._cols[Index] as ExcelTableColumn;
        }
    }

    /// <summary>
    /// Indexer
    /// </summary>
    /// <param name="Name">The name of the table</param>
    /// <returns>The table column. Null if the table name is not found in the collection</returns>
    public ExcelTableColumn this[string Name]
    {
        get
        {
            if (this._colNames.ContainsKey(Name))
            {
                return this._cols[this._colNames[Name]];
            }
            else
            {
                return null;
            }
        }
    }

    IEnumerator<ExcelTableColumn> IEnumerable<ExcelTableColumn>.GetEnumerator() => this._cols.GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => this._cols.GetEnumerator();

    internal string GetUniqueName(string name)
    {
        if (this._colNames.ContainsKey(name))
        {
            int i = 2;

            string? newName;
            do
            {
                newName = name + (i++).ToString(CultureInfo.InvariantCulture);
            } while (this._colNames.ContainsKey(newName));

            return newName;
        }

        return name;
    }

    /// <summary>
    /// Adds one or more columns at the end of the table.
    /// </summary>
    /// <param name="columns">Number of columns to add.</param>
    /// <returns>The added range</returns>
    public ExcelRangeBase Add(int columns = 1) => this.Insert(int.MaxValue, columns);

    /// <summary>
    /// Inserts one or more columns before the specified position in the table.
    /// </summary>
    /// <param name="position">The position in the table where the column will be inserted. 0 will insert the column at the leftmost position. Any value larger than the number of rows in the table will insert a row at the end of the table.</param>
    /// <param name="columns">Number of columns to insert.</param>
    /// <returns>The inserted range</returns>
    public ExcelRangeBase Insert(int position, int columns = 1)
    {
        lock (this.Table)
        {
            ExcelRangeBase? range = this.Table.InsertColumn(position, columns);
            XmlNode refNode;

            if (position >= this._cols.Count)
            {
                refNode = this._cols[this._cols.Count - 1].TopNode;
                position = this._cols.Count;
            }
            else
            {
                refNode = this._cols[position].TopNode;
            }

            for (int i = position; i < position + columns; i++)
            {
                XmlElement? node = this.Table.TableXml.CreateElement("tableColumn", ExcelPackage.schemaMain);

                if (i >= this._cols.Count)
                {
                    _ = refNode.ParentNode.AppendChild(node);
                }
                else
                {
                    _ = refNode.ParentNode.InsertBefore(node, refNode);
                }

                ExcelTableColumn? item = new ExcelTableColumn(this.Table.NameSpaceManager, node, this.Table, i);
                item.Name = this.GetUniqueName($"Column{i + 1}");
                item.Id = this._maxId++;
                this._cols.Insert(i, item);
            }

            for (int i = position; i < this._cols.Count; i++)
            {
                this._cols[i].Position = i;
            }

            this._colNames = this._cols.ToDictionary(x => x.Name, y => y.Position);

            return range;
        }
    }

    /// <summary>
    /// Deletes one or more columns from the specified position in the table.
    /// </summary>
    /// <param name="position">The position in the table where the column will be inserted. 0 will insert the column at the leftmost position. Any value larger than the number of rows in the table will insert a row at the end of the table.</param>
    /// <param name="columns">Number of columns to insert.</param>
    /// <returns>The inserted range</returns>
    public ExcelRangeBase Delete(int position, int columns = 1)
    {
        lock (this.Table)
        {
            if (position < 0)
            {
                throw new ArgumentException("position", "position can't be negative");
            }

            if (columns < 0)
            {
                throw new ArgumentException("columns", "columns can't be negative");
            }

            if (this.Table.Address._toCol < this.Table.Address._fromCol + position + columns - 1)
            {
                throw new InvalidOperationException("Delete will exceed the number of columns in the table");
            }

            for (int i = position + columns - 1; i >= position; i--)
            {
                XmlNode? n = this.Table.Columns[i].TopNode;
                _ = n.ParentNode.RemoveChild(n);
                _ = this.Table.Columns._colNames.Remove(this._cols[i].Name);
                this.Table.Columns._cols.RemoveAt(i);
            }

            for (int i = position; i < this._cols.Count; i++)
            {
                this._cols[i].Position = i;
            }

            this._colNames = this._cols.ToDictionary(x => x.Name, y => y.Position);

            ExcelRangeBase? range = this.Table.DeleteColumn(position, columns);

            return range;
        }
    }
}