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
using System.ComponentModel;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// A collection of pivot table fields
/// </summary>
public class ExcelPivotTableFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableField>
{
    private readonly ExcelPivotTable _table;

    internal ExcelPivotTableFieldCollection(ExcelPivotTable table)
        : base() =>
        this._table = table;

    /// <summary>
    /// Indexer by name
    /// </summary>
    /// <param name="name">The name</param>
    /// <returns>The pivot table field</returns>
    public ExcelPivotTableField this[string name]
    {
        get
        {
            foreach (ExcelPivotTableField? field in this._list)
            {
                if (field.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    return field;
                }
            }

            return null;
        }
    }

    /// <summary>
    /// Returns the date group field.
    /// </summary>
    /// <param name="GroupBy">The type of grouping</param>
    /// <returns>The matching field. If none is found null is returned</returns>
    public ExcelPivotTableField GetDateGroupField(eDateGroupBy GroupBy)
    {
        foreach (ExcelPivotTableField? fld in this._list)
        {
            if (fld.Grouping is ExcelPivotTableFieldDateGroup && ((ExcelPivotTableFieldDateGroup)fld.Grouping).GroupBy == GroupBy)
            {
                return fld;
            }
        }

        return null;
    }

    /// <summary>
    /// Returns the numeric group field.
    /// </summary>
    /// <returns>The matching field. If none is found null is returned</returns>
    public ExcelPivotTableField GetNumericGroupField()
    {
        foreach (ExcelPivotTableField? fld in this._list)
        {
            if (fld.Grouping is ExcelPivotTableFieldNumericGroup)
            {
                return fld;
            }
        }

        return null;
    }

    internal ExcelPivotTableField AddDateGroupField(int index)
    {
        //Pivot field
        XmlElement fieldNode = this.CreateFieldNode(this._table);
        fieldNode.InnerXml = "<items/>";

        ExcelPivotTableField? field = new ExcelPivotTableField(this._table.NameSpaceManager, fieldNode, this._table, this._table.Fields.Count, index);

        this._list.Add(field);

        return field;
    }

    internal ExcelPivotTableField AddField(int index)
    {
        //Pivot field
        XmlElement fieldNode = this.CreateFieldNode(this._table);
        fieldNode.InnerXml = "<items/>";

        ExcelPivotTableField? field = new ExcelPivotTableField(this._table.NameSpaceManager, fieldNode, this._table, this._table.Fields.Count, index);

        this._list.Add(field);

        return field;
    }

    private XmlElement CreateFieldNode(ExcelPivotTable tbl)
    {
        XmlNode? topNode = tbl.PivotTableXml.SelectSingleNode("//d:pivotFields", this._table.NameSpaceManager);
        XmlElement? fieldNode = tbl.PivotTableXml.CreateElement("pivotField", ExcelPackage.schemaMain);
        fieldNode.SetAttribute("compact", "0");
        fieldNode.SetAttribute("outline", "0");
        fieldNode.SetAttribute("showAll", "0");
        fieldNode.SetAttribute("defaultSubtotal", "0");
        _ = topNode.AppendChild(fieldNode);

        return fieldNode;
    }

    /// <summary>
    /// Adds a calculated field to the underlaying pivot table cache. 
    /// </summary>
    /// <param name="name">The unique name of the field</param>
    /// <param name="formula">The formula for the calculated field. 
    /// Note: In formulas you create for calculated fields or calculated items, you can use operators and expressions as you do in other worksheet formulas. You can use constants and refer to data from the  pivot table, but you cannot use cell references or defined names.You cannot use worksheet functions that require cell references or defined names as arguments, and you cannot use array functions.
    /// <seealso cref="ExcelPivotTableCacheField.Formula"/></param>
    /// <returns>The new calculated field</returns>
    public ExcelPivotTableField AddCalculatedField(string name, string formula)
    {
        if (this._list.Exists(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
        {
            throw new InvalidOperationException($"Field with name {name} already exists in the collection");
        }

        PivotTableCacheInternal? cache = this._table.CacheDefinition._cacheReference;
        ExcelPivotTableCacheField? cacheField = cache.AddFormula(name, formula);

        foreach (ExcelPivotTable? pt in cache._pivotTables)
        {
            XmlElement fieldNode = this.CreateFieldNode(pt);
            fieldNode.SetAttribute("dragToPage", "0");
            fieldNode.SetAttribute("dragToCol", "0");
            fieldNode.SetAttribute("dragToRow", "0");
            ExcelPivotTableField? field = new ExcelPivotTableField(this._table.NameSpaceManager, fieldNode, pt, cacheField.Index, 0);
            field._cacheField = cacheField;
            pt.Fields.AddInternal(field);
        }

        return this._table.Fields[cacheField.Index];
    }

    internal void RemoveAt(int index) => this.Remove(this._list[index]);

    internal void Remove(ExcelPivotTableField item)
    {
        _ = item.TopNode.ParentNode.RemoveChild(item.TopNode);
        _ = this._list.Remove(item);
    }
}