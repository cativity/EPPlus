/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/07/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/

using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Sorting;

/// <summary>
/// Preserves the AutoFilter sort state.
/// </summary>
public class SortState : XmlHelper
{
    internal SortState(XmlNamespaceManager nameSpaceManager, XmlNode topNode)
        : base(nameSpaceManager, topNode)
    {
        this._sortConditions = new SortConditionCollection(nameSpaceManager, topNode);
    }

    internal SortState(XmlNamespaceManager nameSpaceManager, ExcelWorksheet worksheet)
        : base(nameSpaceManager, null)
    {
        this.SchemaNodeOrder = worksheet.SchemaNodeOrder;
        this.TopNode = worksheet.WorksheetXml.SelectSingleNode(this._sortStatePath, nameSpaceManager);

        if (this.TopNode == null)
        {
            this.TopNode = this.CreateNode(worksheet.WorksheetXml.DocumentElement, this._sortStatePath);
            XmlAttribute? attr = worksheet.WorksheetXml.CreateAttribute("xmlns:xlrd2");
            attr.Value = ExcelPackage.schemaRichData2;
            this.TopNode.Attributes.Append(attr);
        }
        else
        {
            this.TopNode.RemoveAll();
        }

        this._sortConditions = new SortConditionCollection(nameSpaceManager, this.TopNode);
    }

    internal SortState(XmlNamespaceManager nameSpaceManager, ExcelTable table)
        : base(nameSpaceManager, null)
    {
        this.SchemaNodeOrder = table.SchemaNodeOrder;
        this.TopNode = table.TableXml.SelectSingleNode(this._sortStatePath, nameSpaceManager);

        if (this.TopNode == null)
        {
            this.TopNode = this.CreateNode(table.TableXml.DocumentElement, this._sortStatePath);
            XmlAttribute? attr = table.TableXml.CreateAttribute("xmlns:xlrd2");
            attr.Value = ExcelPackage.schemaRichData2;
            this.TopNode.Attributes.Append(attr);
        }

        this._sortConditions = new SortConditionCollection(nameSpaceManager, this.TopNode);
    }

    private string _sortStatePath = "//d:sortState";
    private string _caseSensitivePath = "@caseSensitive";
    private string _columnSortPath = "@columnSort";
    private string _refPath = "@ref";

    private readonly SortConditionCollection _sortConditions;

    /// <summary>
    /// Removes all sort conditions
    /// </summary>
    public void Clear()
    {
        this._sortConditions.Clear();
    }

    /// <summary>
    /// The preserved sort conditions of the sort state.
    /// </summary>
    public SortConditionCollection SortConditions
    {
        get { return this._sortConditions; }
    }

    /// <summary>
    /// Indicates whether or not the sort is case-sensitive
    /// </summary>
    public bool CaseSensitive
    {
        get { return this.GetXmlNodeBool(this._caseSensitivePath); }
        internal set { this.SetXmlNodeBool(this._caseSensitivePath, value, false); }
    }

    /// <summary>
    /// Indicates whether or not to sort by columns.
    /// </summary>
    public bool ColumnSort
    {
        get { return this.GetXmlNodeBool(this._columnSortPath); }
        internal set { this.SetXmlNodeBool(this._columnSortPath, value, false); }
    }

    /// <summary>
    /// The whole range of data to sort (not only the sort-by column)
    /// </summary>
    public string Ref
    {
        get { return this.GetXmlNodeString(this._refPath); }
        internal set { this.SetXmlNodeString(this._refPath, value); }
    }
}