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
using System.Xml;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.Table;

/// <summary>
/// A collection of table objects
/// </summary>
public class ExcelTableCollection : IEnumerable<ExcelTable>
{
    List<ExcelTable> _tables = new List<ExcelTable>();
    internal Dictionary<string, int> _tableNames = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
    ExcelWorksheet _ws;        
    internal ExcelTableCollection(ExcelWorksheet ws)
    {
        ZipPackage? pck = ws._package.ZipPackage;
        this._ws = ws;
        foreach(XmlElement node in ws.WorksheetXml.SelectNodes("//d:tableParts/d:tablePart", ws.NameSpaceManager))
        {
            ZipPackageRelationship? rel = ws.Part.GetRelationship(node.GetAttribute("id",ExcelPackage.schemaRelationships));
            ExcelTable? tbl = new ExcelTable(rel, ws);
            this._tableNames.Add(tbl.Name, this._tables.Count);
            this._tables.Add(tbl);
        }
    }
    private ExcelTable Add(ExcelTable tbl)
    {
        this._tables.Add(tbl);
        this._tableNames.Add(tbl.Name, this._tables.Count - 1);
        if (tbl.Id >= this._ws.Workbook._nextTableID)
        {
            this._ws.Workbook._nextTableID = tbl.Id + 1;
        }
        return tbl;
    }

    /// <summary>
    /// Create a table on the supplied range
    /// </summary>
    /// <param name="Range">The range address including header and total row</param>
    /// <param name="Name">The name of the table. Must be unique </param>
    /// <returns>The table object</returns>
    public ExcelTable Add(ExcelAddressBase Range, string Name)
    {
        if (Range.WorkSheetName != null && Range.WorkSheetName != this._ws.Name)
        {
            throw new ArgumentException("Range does not belong to a worksheet", "Range");
        }

        if (string.IsNullOrEmpty(Name))
        {
            Name = this.GetNewTableName();
        }
        else
        {
            if (this._ws.Workbook.ExistsTableName(Name))
            {
                throw (new ArgumentException("Tablename is not unique"));
            }
        }

        ValidateName(Name);

        foreach (ExcelTable? t in this._tables)
        {
            if (t.Address.Collide(Range) != ExcelAddressBase.eAddressCollition.No)
            {
                throw (new ArgumentException(string.Format("Table range collides with table {0}", t.Name)));
            }
        }
        foreach (string? mc in this._ws.MergedCells)
        {
            if (mc == null)
            {
                continue; // Issue 780: this happens if a merged cell has been removed
            }

            if (new ExcelAddressBase(mc).Collide(Range) != ExcelAddressBase.eAddressCollition.No)
            {
                throw (new ArgumentException($"Table range collides with merged range {mc}"));
            }
        }

        return this.Add(new ExcelTable(this._ws, Range, Name, this._ws.Workbook._nextTableID));
    }

    private static void ValidateName(string name)
    {
        if (string.IsNullOrEmpty(name.Trim()))
        {
            throw new ArgumentException("Tablename is blank", "Name");
        }

        char c = name[0];
        if (char.IsLetter(c) == false && c != '\\' && c != '_')
        {
            throw new ArgumentException("Tablename start with invalid character", "Name");
        }

        if (!ExcelAddressUtil.IsValidName(name))
        {
            throw (new ArgumentException("Tablename is not valid", "Name"));
        }
    }
    /// <summary>
    /// Delete the table at the specified index
    /// </summary>
    /// <param name="Index">The index</param>
    /// <param name="ClearRange">Clear the rage if set to true</param>
    public void Delete(int Index, bool ClearRange = false)
    {
        this.Delete(this[Index], ClearRange);
    }

    /// <summary>
    /// Delete the table with the specified name
    /// </summary>
    /// <param name="Name">The name of the table to be deleted</param>
    /// <param name="ClearRange">Clear the rage if set to true</param>
    public void Delete(string Name, bool ClearRange = false)
    {
        if (this[Name] == null)
        {
            throw new ArgumentOutOfRangeException(string.Format("Cannot delete non-existant table {0} in sheet {1}.", Name, this._ws.Name));
        }

        this.Delete(this[Name], ClearRange);
    }


    /// <summary>
    /// Delete the table
    /// </summary>
    /// <param name="Table">The table object</param>
    /// <param name="ClearRange">Clear the table range</param>
    public void Delete(ExcelTable Table, bool ClearRange = false)
    {
        if (!this._tables.Contains(Table))
        {
            throw new ArgumentOutOfRangeException("Table", String.Format("Table {0} does not exist in this collection", Table.Name));
        }
        lock (this)
        {
            int tIx = this._tableNames[Table.Name];
            this._tableNames.Remove(Table.Name);
            this._tables.Remove(Table);
            foreach (ExcelWorksheet? sheet in Table.WorkSheet.Workbook.Worksheets)
            {
                if (sheet is ExcelChartsheet)
                {
                    continue;
                }

                foreach (ExcelTable? t in sheet.Tables)
                {
                    if (t.Id > Table.Id)
                    {
                        t.Id--;
                    }
                }
                Table.WorkSheet.Workbook._nextTableID--;
            }
            foreach(string? name in this._tableNames.Keys.ToArray())
            { 
                if(this._tableNames[name] > tIx)
                {
                    this._tableNames[name]--;
                }
            }
            Table.DeleteMe();
            if (ClearRange)
            {
                ExcelRange? range = this._ws.Cells[Table.Address.Address];
                range.Clear();
            }                
        }

    }

    internal string GetNewTableName()
    {
        string name = "Table1";
        int i = 2;
        while (this._ws.Workbook.ExistsTableName(name))
        {
            name = string.Format("Table{0}", i++);
        }
        return name;
    }
    /// <summary>
    /// Number of items in the collection
    /// </summary>
    public int Count
    {
        get
        {
            return this._tables.Count;
        }
    }
    /// <summary>
    /// Get the table object from a range.
    /// </summary>
    /// <param name="Range">The range</param>
    /// <returns>The table. Null if no range matches</returns>
    public static ExcelTable GetFromRange(ExcelRangeBase Range)
    {
        foreach (ExcelTable? tbl in Range.Worksheet.Tables)
        {
            if (tbl.Address._address == Range._address)
            {
                return tbl;
            }
        }
        return null;
    }
    /// <summary>
    /// The table Index. Base 0.
    /// </summary>
    /// <param name="Index"></param>
    /// <returns></returns>
    public ExcelTable this[int Index]
    {
        get
        {
            if (Index < 0 || Index >= this._tables.Count)
            {
                throw (new ArgumentOutOfRangeException("Table index out of range"));
            }
            return this._tables[Index];
        }
    }
    /// <summary>
    /// Indexer
    /// </summary>
    /// <param name="Name">The name of the table</param>
    /// <returns>The table. Null if the table name is not found in the collection</returns>
    public ExcelTable this[string Name]
    {
        get
        {
            if (this._tableNames.ContainsKey(Name))
            {
                return this._tables[this._tableNames[Name]];
            }
            else
            {
                return null;
            }
        }
    }
    /// <summary>
    /// Gets the enumerator for the collection
    /// </summary>
    /// <returns>The enumerator</returns>
    public IEnumerator<ExcelTable> GetEnumerator()
    {
        return this._tables.GetEnumerator();
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return this._tables.GetEnumerator();
    }
}