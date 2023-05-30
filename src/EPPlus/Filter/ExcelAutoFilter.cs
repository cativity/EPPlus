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

using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Xml;
using OfficeOpenXml.Core.CellStore;

namespace OfficeOpenXml.Filter;

/// <summary>
/// Represents an Autofilter for a worksheet or a filter of a table
/// </summary>
public class ExcelAutoFilter : XmlHelper
{
    private const string AutoFilterGuid = "71E0E44A-7884-43F4-9E11-E314B2584A5E";
    private ExcelWorksheet _worksheet;
    //private ExcelTable _table;

    internal ExcelAutoFilter(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelWorksheet worksheet)
        : base(namespaceManager, topNode)
    {
        this._columns = new ExcelFilterColumnCollection(namespaceManager, topNode, this);
        this._worksheet = worksheet;
    }

    internal ExcelAutoFilter(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelTable table)
        : base(namespaceManager, topNode)
    {
        this._columns = new ExcelFilterColumnCollection(namespaceManager, topNode, this);
        this._worksheet = table.WorkSheet;
        //this._table = table;
    }

    internal void Save()
    {
        this.ApplyFilter();

        foreach (ExcelFilterColumn? c in this.Columns)
        {
            c.Save();
        }
    }

    /// <summary>
    /// Applies the filter, hiding rows not matching the filter columns
    /// </summary>
    /// <param name="calculateRange">If true, any formula in the autofilter range will be calculated before the filter is applied.</param>
    public void ApplyFilter(bool calculateRange = false)
    {
        if (calculateRange && this._address != null && ExcelCellBase.IsValidAddress(this._address._address))
        {
            this._worksheet.Cells[this._address._address].Calculate();
        }

        foreach (ExcelFilterColumn? column in this.Columns)
        {
            column.SetFilterValue(this._worksheet, this.Address);
        }

        for (int row = this.Address._fromRow + 1; row <= this._address._toRow; row++)
        {
            RowInternal? rowInternal = ExcelRow.GetRowInternal(this._worksheet, row);
            rowInternal.Hidden = false;

            foreach (ExcelFilterColumn? column in this.Columns)
            {
                ExcelValue value = this._worksheet.GetCoreValueInner(row, this.Address._fromCol + column.Position);
                string? text = ValueToTextHandler.GetFormattedText(value._value, this._worksheet.Workbook, value._styleId, false);

                if (column.Match(value._value, text) == false)
                {
                    rowInternal.Hidden = true;

                    break;
                }
            }
        }
    }

    ExcelAddressBase _address;

    /// <summary>
    /// The range of the autofilter
    /// </summary>
    public ExcelAddressBase Address
    {
        get { return this._address ??= new ExcelAddressBase(this.GetXmlNodeString("@ref")); }
        internal set
        {
            if (value._fromCol != this.Address._fromCol
                || value._toCol != this.Address._toCol
                || value._fromRow != this.Address._fromRow) //Allow different _toRow
            {
                this._columns = new ExcelFilterColumnCollection(this.NameSpaceManager, this.TopNode, this);
            }

            this.SetXmlNodeString("@ref", value.Address);
            this._address = value;
        }
    }

    ExcelFilterColumnCollection _columns;

    /// <summary>
    /// The columns to filter
    /// </summary>
    public ExcelFilterColumnCollection Columns
    {
        get { return this._columns; }
    }
}