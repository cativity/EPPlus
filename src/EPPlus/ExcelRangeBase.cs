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
using System.ComponentModel;
using System.Text;
using System.Data;
using System.Threading;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Style;
using System.Xml;
using System.Drawing;
using System.Globalization;
using System.Collections;
using OfficeOpenXml.Table;
using System.Text.RegularExpressions;
using System.IO;
using System.Linq;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using System.Reflection;
using OfficeOpenXml.Style.XmlAccess;
using System.Security;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using w = System.Windows;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Core;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.Worksheet;
using OfficeOpenXml.ThreadedComments;
using OfficeOpenXml.Sorting;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Export.HtmlExport.Interfaces;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml;

/// <summary>
/// A range of cells 
/// </summary>
public partial class ExcelRangeBase : ExcelAddress, IExcelCell, IDisposable, IEnumerable<ExcelRangeBase>, IEnumerator<ExcelRangeBase>
{
    /// <summary>
    /// Reference to the worksheet
    /// </summary>
    internal protected ExcelWorksheet _worksheet;
    internal ExcelWorkbook _workbook = null;
    private delegate void _changeProp(ExcelRangeBase range, _setValue method, object value);
    private delegate void _setValue(ExcelRangeBase range, object value, int row, int col);
    private _changeProp _changePropMethod;
    private int _styleID;
    #region Constructors
    internal ExcelRangeBase(ExcelWorksheet xlWorksheet)
    {
        this.Init(xlWorksheet);
        this._ws = this._worksheet.Name;
        this._workbook = this._worksheet.Workbook;
        this.SetDelegate();
    }

    internal ExcelRangeBase(ExcelWorksheet xlWorksheet, string address) :
        base(xlWorksheet == null ? "" : xlWorksheet.Name, address)
    {
        this.Init(xlWorksheet);
        this._workbook = this._worksheet.Workbook;
        this.SetRCFromTable(this._worksheet._package, null);
        if (string.IsNullOrEmpty(this._ws))
        {
            this._ws = this._worksheet == null ? "" : this._worksheet.Name;
        }

        this.SetDelegate();
    }
    internal ExcelRangeBase(ExcelWorkbook wb, ExcelWorksheet xlWorksheet, string address, bool isName) :
        base(xlWorksheet == null ? "" : xlWorksheet.Name, address, isName)
    {
        this.Init(xlWorksheet);
        this.SetRCFromTable(wb._package, null);
        this._workbook = wb;
        if (string.IsNullOrEmpty(this._ws))
        {
            this._ws = (xlWorksheet == null ? null : xlWorksheet.Name);
        }

        this.SetDelegate();
    }
    #endregion
    private void Init(ExcelWorksheet xlWorksheet)
    {
        this._worksheet = xlWorksheet;            
    }

    /// <summary>
    /// On change address handler
    /// </summary>
    protected internal override void ChangeAddress()
    {
        if (this.Table != null)
        {
            this.SetRCFromTable(this._workbook._package, null);
        }
        if (string.IsNullOrEmpty(this._ws) == false && (this._worksheet == null || !this._worksheet.Name.Equals(this._ws, StringComparison.OrdinalIgnoreCase)))
        {
            this._worksheet = this._workbook.Worksheets[this._ws];
        }

        this.SetDelegate();
    }
    #region Set Value Delegates        
    private static _changeProp _setUnknownProp = SetUnknown;
    private static _changeProp _setSingleProp = SetSingle;
    private static _changeProp _setRangeProp = SetRange;
    private static _changeProp _setMultiProp = SetMultiRange;
    private void SetDelegate()
    {
        if (this._fromRow == -1)
        {
            this._changePropMethod = SetUnknown;
        }
        //Single cell
        else if (this._fromRow == this._toRow && this._fromCol == this._toCol && this.Addresses == null)
        {
            this._changePropMethod = SetSingle;
        }
        //Range (ex A1:A2)
        else if (this.Addresses == null)
        {
            this._changePropMethod = SetRange;
        }
        //Multi Range (ex A1:A2,C1:C2)
        else
        {
            this._changePropMethod = SetMultiRange;
        }
    }
    /// <summary>
    /// We dont know the address yet. Set the delegate first time a property is set.
    /// </summary>
    /// <param name="range"></param>
    /// <param name="valueMethod"></param>
    /// <param name="value"></param>
    private static void SetUnknown(ExcelRangeBase range, _setValue valueMethod, object value)
    {
        //Address is not set use, selected range
        if (range._fromRow == -1)
        {
            range.SetToSelectedRange();
        }
        range.SetDelegate();
        range._changePropMethod(range, valueMethod, value);
    }
    /// <summary>
    /// Set a single cell
    /// </summary>
    /// <param name="range"></param>
    /// <param name="valueMethod"></param>
    /// <param name="value"></param>
    private static void SetSingle(ExcelRangeBase range, _setValue valueMethod, object value)
    {
        valueMethod(range, value, range._fromRow, range._fromCol);
    }
    /// <summary>
    /// Set a range
    /// </summary>
    /// <param name="range"></param>
    /// <param name="valueMethod"></param>
    /// <param name="value"></param>
    private static void SetRange(ExcelRangeBase range, _setValue valueMethod, object value)
    {
        range.SetValueAddress(range, valueMethod, value);
    }
    /// <summary>
    /// Set a multirange (A1:A2,C1:C2)
    /// </summary>
    /// <param name="range"></param>
    /// <param name="valueMethod"></param>
    /// <param name="value"></param>
    private static void SetMultiRange(ExcelRangeBase range, _setValue valueMethod, object value)
    {
        //range.SetValueAddress(range, valueMethod, value);
        foreach (ExcelAddressBase? address in range.Addresses)
        {
            range.SetValueAddress(address, valueMethod, value);
        }
    }
    /// <summary>
    /// Set the property for an address
    /// </summary>
    /// <param name="address"></param>
    /// <param name="valueMethod"></param>
    /// <param name="value"></param>
    private void SetValueAddress(ExcelAddressBase address, _setValue valueMethod, object value)
    {
        this.IsRangeValid("");
        if (this._fromRow == 1 && this._fromCol == 1 && this._toRow == ExcelPackage.MaxRows && this._toCol == ExcelPackage.MaxColumns)  //Full sheet (ex ws.Cells.Value=0). Set value for A1 only to avoid hanging 
        {
            throw (new ArgumentException("Can't reference all cells. Please use the indexer to set the range"));
        }
        else
        {
            if (value is object[,] && (valueMethod == Set_Value || valueMethod == Set_StyleID))
            {
                // only simple set value is supported for bulk copy
                this._worksheet.SetRangeValueInner(address.Start.Row, address.Start.Column, address.End.Row, address.End.Column, (object[,])value, false);
            }
            else
            {
                if (valueMethod != Set_IsRichText)
                {
                    this.DeleteMe(address, false, false, true, true, false, false, false, false, false);   //Clear the range before overwriting, but not merged cells.
                }

                for (int col = address.Start.Column; col <= address.End.Column; col++)
                {
                    for (int row = address.Start.Row; row <= address.End.Row; row++)
                    {
                        valueMethod(this, value, row, col);
                    }
                }
            }
        }
    }
    #endregion
    #region Set property methods
    private static _setValue _setStyleIdDelegate = Set_StyleID;
    private static _setValue _setValueDelegate = Set_Value;
    private static _setValue _setHyperLinkDelegate = Set_HyperLink;
    private static _setValue _setIsRichTextDelegate = Set_IsRichText;
    private static _setValue _setExistsCommentDelegate = Exists_Comment;
    private static _setValue _setCommentDelegate = Set_Comment;
    private static _setValue _setExistsThreadedCommentDelegate = Exists_ThreadedComment;
    private static _setValue _setThreadedCommentDelegate = Set_ThreadedComment;

    private static void Set_StyleID(ExcelRangeBase range, object value, int row, int col)
    {
        range._worksheet.SetStyleInner(row, col, (int)value);
    }
    private static void Set_StyleName(ExcelRangeBase range, object value, int row, int col)
    {
        range._worksheet.SetStyleInner(row, col, range._styleID);
    }
    private static void Set_Value(ExcelRangeBase range, object value, int row, int col)
    {
        object? sfi = range._worksheet._formulas.GetValue(row, col);
        if (sfi is int)
        {
            range.SplitFormulas(range._worksheet.Cells[row, col]);
        }
        if (sfi != null)
        {
            range._worksheet._formulas.SetValue(row, col, string.Empty);
        }

        range._worksheet.SetValueInner(row, col, value);
        range._worksheet._flags.Clear(row, col, 1, 1);
        range._worksheet._metadataStore.Clear(row, col, 1, 1);
    }
    private static void Set_Formula(ExcelRangeBase range, object value, int row, int col)
    {
        object? f = range._worksheet._formulas.GetValue(row, col);
        if (f is int && (int)f >= 0)
        {
            range.SplitFormulas(range._worksheet.Cells[row, col]);
        }

        string formula = (value == null ? string.Empty : value.ToString());
        if (formula == string.Empty)
        {
            range._worksheet._formulas.SetValue(row, col, string.Empty);
        }
        else
        {
            if (formula[0] == '=')
            {
                formula = formula.Substring(1, formula.Length - 1); // remove any starting equalsign.
            }

            range._worksheet._formulas.SetValue(row, col, formula);
            range._worksheet.SetValueInner(row, col, null);
        }
    }
    /// <summary>
    /// Handles shared formulas
    /// </summary>
    /// <param name="range">The range</param>
    /// <param name="value">The  formula</param>
    /// <param name="address">The address of the formula</param>
    /// <param name="IsArray">If the forumla is an array formula.</param>
    private static void Set_SharedFormula(ExcelRangeBase range, string value, ExcelAddressBase address, bool IsArray)
    {
        if (range._fromRow == 1 && range._fromCol == 1 && range._toRow == ExcelPackage.MaxRows && range._toCol == ExcelPackage.MaxColumns)  //Full sheet (ex ws.Cells.Value=0). Set value for A1 only to avoid hanging 
        {
            throw (new InvalidOperationException("Can't set a formula for the entire worksheet"));
        }
        else if (address.Start.Row == address.End.Row && address.Start.Column == address.End.Column && !IsArray)             //is it really a shared formula? Arrayformulas can be one cell only
        {
            //Nope, single cell. Set the formula
            Set_Formula(range, value, address.Start.Row, address.Start.Column);
            return;
        }

        range.CheckAndSplitSharedFormula(address);
        ExcelWorksheet.Formulas f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default);
        f.Formula = value;
        f.Index = range._worksheet.GetMaxShareFunctionIndex(IsArray);
        f.Address = address.FirstAddress;
        f.StartCol = address.Start.Column;
        f.StartRow = address.Start.Row;
        f.FormulaType = IsArray ? ExcelWorksheet.FormulaType.Array : ExcelWorksheet.FormulaType.Shared;

        range._worksheet._sharedFormulas.Add(f.Index, f);

        for (int col = address.Start.Column; col <= address.End.Column; col++)
        {
            for (int row = address.Start.Row; row <= address.End.Row; row++)
            {
                range._worksheet._formulas.SetValue(row, col, f.Index);
                range._worksheet._flags.SetFlagValue(row, col, true, CellFlags.ArrayFormula);
                range._worksheet.SetValueInner(row, col, null);
            }
        }
    }

    private static void Set_HyperLink(ExcelRangeBase range, object value, int row, int col)
    {
        if (value is Uri)
        {
            range._worksheet._hyperLinks.SetValue(row, col, (Uri)value);

            if (value is ExcelHyperLink hl)
            {                    
                if (string.IsNullOrEmpty(hl.Display))
                {
                    object? v = range._worksheet.GetValueInner(row, col);
                    if(v == null)
                    {
                        range._worksheet.SetValueInner(row, col, hl.ReferenceAddress);
                    }
                }
                else
                {
                    range._worksheet.SetValueInner(row, col, hl.Display);
                }
            }
            else
            {
                object? v = range._worksheet.GetValueInner(row, col);
                if (v == null || v.ToString() == "")
                {
                    range._worksheet.SetValueInner(row, col, ((Uri)value).OriginalString);
                }
            }
        }
        else
        {
            range._worksheet._hyperLinks.SetValue(row, col, null);
            range._worksheet.SetValueInner(row, col, null);
        }
    }
    private static void Set_IsRichText(ExcelRangeBase range, object value, int row, int col)
    {
        bool b = (bool)value;
        ExcelWorksheet? ws = range.Worksheet;
        bool isRT = ws._flags.GetFlagValue(row, col, CellFlags.RichText);
        if (isRT != b)
        {
            ExcelRichTextCollection? rt = ws.GetRichText(row, col, ws.Cells[row, col]);
            if (b)
            {
                rt.Text = ValueToTextHandler.GetFormattedText(ws.GetValue(row, col), ws.Workbook, ws.GetStyleInner(row, col), false);
            }
            else
            {
                Set_Value(range, rt.Text, row, col);
            }

            range._worksheet._flags.SetFlagValue(row, col, (bool)value, CellFlags.RichText);
        }
    }
    private static void Exists_Comment(ExcelRangeBase range, object value, int row, int col)
    {
        Exists_ThreadedComment(range, value, row, col);
        if (range._worksheet._commentsStore.Exists(row, col))
        {
            throw (new InvalidOperationException(string.Format("Cell {0} already contain a comment.", new ExcelCellAddress(row, col).Address)));
        }

    }
    private static void Set_Comment(ExcelRangeBase range, object value, int row, int col)
    {
        string[] v = (string[])value;
        range._worksheet.Comments.Add(new ExcelRangeBase(range._worksheet, GetAddress(row, col)), v[0], v[1]);
    }
    private static void Exists_ThreadedComment(ExcelRangeBase range, object value, int row, int col)
    {
        if (range._worksheet._threadedCommentsStore.Exists(row, col))
        {
            throw (new InvalidOperationException(string.Format("Cell {0} already contain a threaded comment.", new ExcelCellAddress(row, col).Address)));
        }

    }
    private static void Set_ThreadedComment(ExcelRangeBase range, object value, int row, int col)
    {
        range._worksheet.ThreadedComments.Add(GetAddress(row, col));
    }

    #endregion
    internal void SetToSelectedRange()
    {
        if (this._worksheet.View.SelectedRange == "")
        {
            this.Address = "A1";
        }
        else
        {
            this.Address = this._worksheet.View.SelectedRange;
        }
    }
    private void IsRangeValid(string type)
    {
        if (this._fromRow <= 0)
        {
            if (this._address == "")
            {
                this.SetToSelectedRange();
            }
            else
            {
                if (type == "")
                {
                    throw (new InvalidOperationException(string.Format("Range is not valid for this operation: {0}", this._address)));
                }
                else
                {
                    throw (new InvalidOperationException(string.Format("Range is not valid for {0} : {1}", type, this._address)));
                }
            }
        }
    }
    #region Public Properties
    /// <summary>
    /// The style object for the range.
    /// </summary>
    public ExcelStyle Style
    {
        get
        {
            this.IsRangeValid("styling");
            int s = 0;
            if (!this._worksheet.ExistsStyleInner(this._fromRow, this._fromCol, ref s)) //Cell exists
            {
                if (!this._worksheet.ExistsStyleInner(this._fromRow, 0, ref s)) //No, check Row style
                {
                    ExcelColumn? c = this.Worksheet.GetColumn(this._fromCol);
                    if (c == null)
                    {
                        s = 0;
                    }
                    else
                    {
                        s = c.StyleID;
                    }
                }
            }
            return this._worksheet.Workbook.Styles.GetStyleObject(s, this._worksheet.PositionId, this.Address);
        }
    }
    /// <summary>
    /// The named style
    /// </summary>
    public string StyleName
    {
        get
        {
            this.IsRangeValid("styling");
            int xfId;
            if (this._fromRow == 1 && this._toRow == ExcelPackage.MaxRows)
            {
                xfId = this.GetColumnStyle(this._fromCol);
            }
            else if (this._fromCol == 1 && this._toCol == ExcelPackage.MaxColumns)
            {
                xfId = 0;
                if (!this._worksheet.ExistsStyleInner(this._fromRow, 0, ref xfId))
                {
                    xfId = this.GetColumnStyle(this._fromCol);
                }
            }
            else
            {
                xfId = 0;
                if (!this._worksheet.ExistsStyleInner(this._fromRow, this._fromCol, ref xfId))
                {
                    if (!this._worksheet.ExistsStyleInner(this._fromRow, 0, ref xfId))
                    {
                        xfId = this.GetColumnStyle(this._fromCol);
                    }
                }
            }
            int nsID;
            if (xfId <= 0)
            {
                nsID = this.Style.Styles.CellXfs[0].XfId;
            }
            else
            {
                nsID = this.Style.Styles.CellXfs[xfId].XfId;
            }
            foreach (ExcelNamedStyleXml? ns in this.Style.Styles.NamedStyles)
            {
                if (ns.StyleXfId == nsID)
                {
                    return ns.Name;
                }
            }

            return "";
        }
        set
        {
            this._styleID = this._worksheet.Workbook.Styles.GetStyleIdFromName(value);
            int col = this._fromCol;
            if (this._fromRow == 1 && this._toRow == ExcelPackage.MaxRows)    //Full column
            {
                ExcelColumn column;
                object? c = this._worksheet.GetValue(0, this._fromCol);
                if (c == null)
                {
                    column = this._worksheet.Column(this._fromCol);
                }
                else
                {
                    column = (ExcelColumn)c;
                }

                column.StyleName = value;
                column.StyleID = this._styleID;

                CellStoreEnumerator<ExcelValue>? cols = new CellStoreEnumerator<ExcelValue>(this._worksheet._values, 0, this._fromCol + 1, 0, this._toCol);
                if (cols.Next())
                {
                    col = this._fromCol;
                    while (column.ColumnMin <= this._toCol)
                    {
                        if (column.ColumnMax > this._toCol)
                        {
                            ExcelColumn? newCol = this._worksheet.CopyColumn(column, this._toCol + 1, column.ColumnMax);
                            column.ColumnMax = this._toCol;
                        }

                        column._styleName = value;
                        column.StyleID = this._styleID;

                        if (cols.Value._value == null)
                        {
                            break;
                        }
                        else
                        {
                            ExcelColumn? nextCol = (ExcelColumn)cols.Value._value;
                            if (column.ColumnMax < nextCol.ColumnMax - 1)
                            {
                                column.ColumnMax = nextCol.ColumnMax - 1;
                            }
                            column = nextCol;
                            cols.Next();
                        }
                    }
                }
                if (column.ColumnMax < this._toCol)
                {
                    column.ColumnMax = this._toCol;
                }

                if (this._fromCol == 1 && this._toCol == ExcelPackage.MaxColumns) //FullRow
                {
                    CellStoreEnumerator<ExcelValue>? rows = new CellStoreEnumerator<ExcelValue>(this._worksheet._values, 1, 0, ExcelPackage.MaxRows, 0);
                    rows.Next();
                    while (rows.Value._value != null)
                    {
                        this._worksheet.SetStyleInner(rows.Row, 0, this._styleID);
                        if (!rows.Next())
                        {
                            break;
                        }
                    }
                }
            }
            else if (this._fromCol == 1 && this._toCol == ExcelPackage.MaxColumns) //FullRow
            {
                for (int r = this._fromRow; r <= this._toRow; r++)
                {
                    this._worksheet.Row(r)._styleName = value;
                    this._worksheet.Row(r).StyleID = this._styleID;
                }
            }

            if (!((this._fromRow == 1 && this._toRow == ExcelPackage.MaxRows) || (this._fromCol == 1 && this._toCol == ExcelPackage.MaxColumns))) //Cell specific
            {
                for (int c = this._fromCol; c <= this._toCol; c++)
                {
                    for (int r = this._fromRow; r <= this._toRow; r++)
                    {
                        this._worksheet.SetStyleInner(r, c, this._styleID);
                    }
                }
            }
            else //Only set name on created cells. (uncreated cells is set on full row or full column).
            {
                CellStoreEnumerator<ExcelValue>? cells = new CellStoreEnumerator<ExcelValue>(this._worksheet._values, this._fromRow, this._fromCol, this._toRow, this._toCol);
                while (cells.Next())
                {
                    this._worksheet.SetStyleInner(cells.Row, cells.Column, this._styleID);
                }
            }
        }
    }

    private int GetColumnStyle(int col)
    {
        object c = null;
        if (this._worksheet.ExistsValueInner(0, col, ref c))
        {
            return (c as ExcelColumn).StyleID;
        }
        else
        {
            int row = 0;
            if (this._worksheet._values.PrevCell(ref row, ref col))
            {
                ExcelValue v = this._worksheet.GetCoreValueInner(row, col);
                ExcelColumn? column = (ExcelColumn)v._value;
                if (column.ColumnMax >= col)
                {
                    return v._styleId;
                }
            }
        }
        return 0;
    }
    /// <summary>
    /// The style ID. 
    /// It is not recomended to use this one. Use Named styles as an alternative.
    /// If you do, make sure that you use the Style.UpdateXml() method to update any new styles added to the workbook.
    /// </summary>
    public int StyleID
    {
        get
        {
            int s = 0;
            if (!this._worksheet.ExistsStyleInner(this._fromRow, this._fromCol, ref s))
            {
                if (!this._worksheet.ExistsStyleInner(this._fromRow, 0, ref s))
                {
                    s = this._worksheet.GetStyleInner(0, this._fromCol);
                }
            }
            return s;
        }
        set
        {
            this._changePropMethod(this, _setStyleIdDelegate, value);
        }
    }
    /// <summary>
    /// Set the range to a specific value
    /// </summary>
    public object Value
    {
        get
        {
            if (this.IsName)
            {
                if (this._worksheet == null)
                {
                    return this._workbook._names[this._address].NameValue;
                }
                else
                {
                    return this._worksheet.Names[this._address].NameValue;
                }
            }
            else
            {
                if (this._fromRow == this._toRow && this._fromCol == this._toCol)
                {
                    return this._worksheet.GetValue(this._fromRow, this._fromCol);
                }
                else
                {
                    return this.GetValueArray();
                }
            }
        }
        set
        {
            if (this.IsName)
            {
                if (this._worksheet == null)
                {
                    this._workbook._names[this._address].NameValue = value;
                }
                else
                {
                    this._worksheet.Names[this._address].NameValue = value;
                }
            }
            else
            {
                this._changePropMethod(this, _setValueDelegate, value);
            }
        }
    }
    /// <summary>
    /// Sets the range to an Error value
    /// </summary>
    /// <param name="errorType">The type of error</param>
    public void SetErrorValue(eErrorType errorType)
    {
        this.Value = ExcelErrorValue.Create(errorType);
    }
    private object GetValueArray()
    {
        ExcelAddressBase addr;
        if (this._fromRow == 1 && this._fromCol == 1 && this._toRow == ExcelPackage.MaxRows && this._toCol == ExcelPackage.MaxColumns)
        {
            addr = this._worksheet.Dimension;
            if (addr == null)
            {
                return null;
            }
        }
        else
        {
            addr = this;
        }
        object[,] v = new object[addr._toRow - addr._fromRow + 1, addr._toCol - addr._fromCol + 1];

        for (int col = addr._fromCol; col <= addr._toCol; col++)
        {
            for (int row = addr._fromRow; row <= addr._toRow; row++)
            {
                object o = null;
                if (this._worksheet.ExistsValueInner(row, col, ref o))
                {
                    if (this._worksheet._flags.GetFlagValue(row, col, CellFlags.RichText))
                    {
                        v[row - addr._fromRow, col - addr._fromCol] = this._worksheet.GetRichText(row, col, this).Text;
                    }
                    else
                    {
                        v[row - addr._fromRow, col - addr._fromCol] = o;
                    }
                }
            }
        }
        return v;
    }
    private ExcelAddressBase GetAddressDim(ExcelRangeBase addr)
    {
        ExcelAddressBase? d = this._worksheet.Dimension;
        int fromRow = addr._fromRow < d._fromRow ? d._fromRow : addr._fromRow;
        int fromCol = addr._fromCol < d._fromCol ? d._fromCol : addr._fromCol;

        int toRow = addr._toRow > d._toRow ? d._toRow : addr._toRow;
        int toCol = addr._toCol > d._toCol ? d._toCol : addr._toCol;

        if (addr._fromRow == fromRow && addr._fromCol == fromCol && addr._toRow == toRow && addr._toCol == this._toCol)
        {
            return addr;
        }
        else
        {
            if (this._fromRow > this._toRow || this._fromCol > this._toCol)
            {
                return null;
            }
            else
            {
                return new ExcelAddressBase(fromRow, fromCol, toRow, toCol);
            }
        }
    }

    private object GetSingleValue()
    {
        if (this.IsRichText)
        {
            return this.RichText.Text;
        }
        else
        {
            return this._worksheet.GetValueInner(this._fromRow, this._fromCol);
        }
    }
    /// <summary>
    /// Returns the formatted value.
    /// </summary>
    public string Text
    {
        get
        {
            if (this.IsSingleCell || this.IsName)
            {
                return ValueToTextHandler.GetFormattedText(this.Value, this._workbook, this.StyleID, false);
            }
            else
            {
                return ValueToTextHandler.GetFormattedText(this._worksheet.GetValue(this._fromRow, this._fromCol), this._workbook, this.StyleID, false);
            }
        }
    }
    /// <summary>
    /// Set the column width from the content of the range. Columns outside of the worksheets dimension are ignored.
    /// The minimum width is the value of the ExcelWorksheet.defaultColumnWidth property.
    /// </summary>
    /// <remarks>
    /// Cells containing formulas must be calculated before autofit is called.
    /// Wrapped and merged cells are also ignored.
    /// </remarks>
    public void AutoFitColumns()
    {
        this.AutoFitColumns(this._worksheet.DefaultColWidth);
    }
    /// <summary>
    /// Set the column width from the content of the range. Columns outside of the worksheets dimension are ignored.
    /// </summary>
    /// <remarks>
    /// This method will not work if you run in an environment that does not support GDI.
    /// Cells containing formulas are ignored if no calculation is made.
    /// Wrapped and merged cells are also ignored.
    /// </remarks>
    /// <param name="MinimumWidth">Minimum column width</param>
    public void AutoFitColumns(double MinimumWidth)
    {
        this.AutoFitColumns(MinimumWidth, double.MaxValue);
    }

    /// <summary>
    /// Set the column width from the content of the range. Columns outside of the worksheets dimension are ignored.
    /// </summary>
    /// <remarks>
    /// This method will not work if you run in an environment that does not support GDI.
    /// Cells containing formulas are ignored if no calculation is made.
    /// Wrapped and merged cells are also ignored.
    /// </remarks>        
    /// <param name="MinimumWidth">Minimum column width</param>
    /// <param name="MaximumWidth">Maximum column width</param>
    public void AutoFitColumns(double MinimumWidth, double MaximumWidth)
    {
#if (Core)
        //var af = new AutofitHelperSkia(this);
        //af.AutofitColumn(MinimumWidth, MaximumWidth);
        AutofitHelper? af = new AutofitHelper(this);
        af.AutofitColumn(MinimumWidth, MaximumWidth);
#else
            var af = new AutofitHelper(this);
            af.AutofitColumn(MinimumWidth, MaximumWidth);
#endif
    }
    internal string TextForWidth
    {
        get
        {
            return ValueToTextHandler.GetFormattedText(this.Value, this._workbook, this.StyleID, true);
        }
    }

    /// <summary>
    /// Gets or sets a formula for a range.
    /// </summary>
    public string Formula
    {
        get
        {
            if (this.IsName)
            {
                if (this._worksheet == null)
                {
                    return this._workbook._names[this._address].NameFormula;
                }
                else
                {
                    return this._worksheet.Names[this._address].NameFormula;
                }
            }
            else
            {
                return this._worksheet.GetFormula(this._fromRow, this._fromCol);
            }
        }
        set
        {
            if (this.IsName)
            {
                if (this._worksheet == null)
                {
                    this._workbook._names[this._address].NameFormula = value;
                }
                else
                {
                    this._worksheet.Names[this._address].NameFormula = value;
                }
            }
            else
            {
                if (value == null || value.Trim() == "")
                {
                    this.Value = null;
                    return;
                }
                if (this._fromRow == this._toRow && this._fromCol == this._toCol)
                {
                    Set_Formula(this, value, this._fromRow, this._fromCol);
                }
                else if (this.HasOffSheetReference(value))
                {
                    this.Set_Formula_Range(this, value);
                }
                else
                {
                    Set_SharedFormula(this, value, this, false);
                    if (this.Addresses != null)
                    {
                        foreach (ExcelAddressBase? address in this.Addresses)
                        {
                            Set_SharedFormula(this, value, address, false);
                        }
                    }
                }

                this.ClearTableFormulas();
            }
        }
    }

    private void Set_Formula_Range(ExcelRangeBase range, string formula)
    {
        if (formula[0] == '=')
        {
            formula = formula.Substring(1); // remove any starting equalsign.
        }

        range.CheckAndSplitSharedFormula(range);

        ExcelWorksheet.Formulas f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default);
        f.Formula = formula;
        f.Address = range.FirstAddress;
        f.StartCol = range.Start.Column;
        f.StartRow = range.Start.Row;

        if (range.Addresses == null)
        {
            this.SetFormulaAddress(range, range, f);
        }
        else
        {
            foreach (ExcelAddressBase? address in range.Addresses)
            {
                this.SetFormulaAddress(range, address, f);
            }
        }
    }

    private void SetFormulaAddress(ExcelRangeBase range, ExcelAddressBase address, ExcelWorksheet.Formulas f)
    {
        for (int row = address._fromRow; row <= address._toRow; row++)
        {
            for (int col = address._fromCol; col <= address._toCol; col++)
            {
                if (string.IsNullOrEmpty(f.Formula))
                {
                    range._worksheet._formulas.SetValue(row, col, string.Empty);
                }
                else
                {
                    range._worksheet._formulas.SetValue(row, col, f.GetFormula(row, col, this.WorkSheetName));
                    range._worksheet.SetValueInner(row, col, null);
                }
            }
        }
    }

    private bool HasOffSheetReference(string value)
    {
        ISourceCodeTokenizer? tokenizer = SourceCodeTokenizer.Default;
        IEnumerable<Token>? tokens = tokenizer.Tokenize(value, this.WorkSheetName);
        foreach (Token t in tokens)
        {
            if (t.TokenTypeIsSet(TokenType.ExcelAddress))
            {
                ExcelAddressBase? a = new ExcelAddressBase(t.Value);
                if (string.IsNullOrEmpty(a.WorkSheetName) == false && a.WorkSheetName.Equals(this.WorkSheetName) == false)
                {
                    return true;
                }
            }
        }
        return false;
    }

    /// <summary>
    /// Gets or Set a formula in R1C1 format.
    /// </summary>
    public string FormulaR1C1
    {
        get
        {
            this.IsRangeValid("FormulaR1C1");
            return this._worksheet.GetFormulaR1C1(this._fromRow, this._fromCol);
        }
        set
        {
            this.IsRangeValid("FormulaR1C1");
            if (value.Length > 0 && value[0] == '=')
            {
                value = value.Substring(1, value.Length - 1); // remove any starting equalsign.
            }

            if (value == null || value.Trim() == "")
            {
                //Set the cells to null
                this.Value = null;
            }
            else
            {
                string? formula = TranslateFromR1C1(value, this._fromRow, this._fromCol);
                if (this._fromRow == this._toRow && this._fromCol == this._toCol)
                {
                    Set_Formula(this, formula, this._fromRow, this._fromCol);
                }
                else if (this.HasOffSheetReference(formula))
                {
                    this.Set_Formula_Range(this, formula);
                }
                else
                {
                    Set_SharedFormula(this, formula, this, false);
                    if (this.Addresses != null)
                    {
                        foreach (ExcelAddressBase? address in this.Addresses)
                        {
                            formula = TranslateFromR1C1(value, address._fromRow, address._fromCol);
                            Set_SharedFormula(this, formula, address, false);
                        }
                    }
                }

                this.ClearTableFormulas();
            }
        }
    }

    /// <summary>
    /// Creates an <see cref="IExcelHtmlRangeExporter"/> for html export of this range.
    /// </summary>
    /// <returns>A html exporter</returns>
    public IExcelHtmlRangeExporter CreateHtmlExporter()
    {
        return new Export.HtmlExport.Exporters.ExcelHtmlRangeExporter(this);
    }

    //public ExcelHtmlRangeExporter CreateHtmlExporter()
    //{
    //    return new ExcelHtmlRangeExporter(this);
    //}
    /// <summary>
    /// Set the Hyperlink property for a range of cells
    /// </summary>
    public Uri Hyperlink
    {
        get
        {
            this.IsRangeValid("formulaR1C1");
            return this._worksheet._hyperLinks.GetValue(this._fromRow, this._fromCol);
        }
        set
        {
            this._changePropMethod(this, _setHyperLinkDelegate, value);
        }
    }
    /// <summary>
    /// Sets the hyperlink property
    /// </summary>
    /// <param name="uri">The URI to set</param>
    public void SetHyperlink(Uri uri)
    {
        this.Hyperlink = uri;
    }
    /// <summary>
    /// Sets the Hyperlink property using the ExcelHyperLink class.
    /// </summary>
    /// <param name="uri">The <see cref="ExcelHyperLink"/> uri to set</param>
    public void SetHyperlink(ExcelHyperLink uri)
    {
        this.Hyperlink = uri;
    }
    /// <summary>
    /// Sets the Hyperlink property to an url within the workbook.
    /// </summary>
    /// <param name="range">A reference within the same workbook</param>
    /// <param name="display">The displayed text in the cell. If display is null or empty, the address of the range will be set.</param>f
    public void SetHyperlink(ExcelRange range, string display)
    {
        if (string.IsNullOrEmpty(display))
        {
            display = range.Address;
        }

        this.SetHyperlinkLocal(range, display);
    }
    /// <summary>
    /// Sets the Hyperlink property to an url within the workbook. The hyperlink will display the value of the cell.
    /// </summary>
    /// <param name="range">A reference within the same workbook</param>
    public void SetHyperlink(ExcelRange range)
    {
        this.SetHyperlinkLocal(range, null);
    }
    private void SetHyperlinkLocal(ExcelRange range, string display)
    {
        if (range == null)
        {
            throw (new ArgumentNullException("The range must not be null.", nameof(range)));
        }
        if (range.Worksheet.Workbook != this.Worksheet.Workbook)
        {
            throw (new ArgumentException("The range must be within this package.", nameof(range)));
        }
        if (string.IsNullOrEmpty(range.WorkSheetName) || range.WorkSheetName.Equals(this.WorkSheetName ?? "", StringComparison.OrdinalIgnoreCase))
        {
            this.Hyperlink = new ExcelHyperLink(range.Address, display);
        }
        else
        {
            this.Hyperlink = new ExcelHyperLink(range.FullAddress, display);
        }
    }
    /// <summary>
    /// If the cells in the range are merged.
    /// </summary>
    public bool Merge
    {
        get
        {
            this.IsRangeValid("merging");
            for (int col = this._fromCol; col <= this._toCol; col++)
            {
                for (int row = this._fromRow; row <= this._toRow; row++)
                {
                    if (this._worksheet.MergedCells[row, col] == null)
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        set
        {
            this.IsRangeValid("merging");
            this.ValidateMergePossible();
            this._worksheet.MergedCells.Clear(this);
            if (value)
            {
                this._worksheet.MergedCells.Add(new ExcelAddressBase(this.FirstAddress), true);
                if (this.Addresses != null)
                {
                    foreach (ExcelAddressBase? address in this.Addresses)
                    {
                        this._worksheet.MergedCells.Clear(address); //Fixes issue 15482
                        this._worksheet.MergedCells.Add(address, true);
                    }
                }
            }
            else
            {
                if (this.Addresses != null)
                {
                    foreach (ExcelAddressBase? address in this.Addresses)
                    {
                        this._worksheet.MergedCells.Clear(address);
                    }
                }

            }
        }
    }

    private void ValidateMergePossible()
    {
        foreach (ExcelTable? t in this._worksheet.Tables)
        {
            if (this.Collide(t.Address) != eAddressCollition.No)
            {
                throw (new InvalidOperationException($"Cant merge range. The merge is within table {t.Name}"));
            }
        }
    }

    /// <summary>
    /// Set an autofilter for the range
    /// </summary>
    public bool AutoFilter
    {
        get
        {
            this.IsRangeValid("autofilter");
            ExcelAddressBase address = this._worksheet.AutoFilterAddress;
            if (address == null)
            {
                return false;
            }

            if (this._fromRow >= address.Start.Row
                && this._toRow <= address.End.Row
                && this._fromCol >= address.Start.Column
                && this._toCol <= address.End.Column)
            {
                return true;
            }
            return false;
        }
        set
        {
            this.IsRangeValid("autofilter");
            if (this._worksheet.AutoFilterAddress != null)
            {
                eAddressCollition c = this.Collide(this._worksheet.AutoFilterAddress);
                if (value == false && (c == eAddressCollition.Partly || c == eAddressCollition.No))
                {
                    throw (new InvalidOperationException("Can't remove Autofilter. The current autofilter does not match selected range."));
                }
            }
            if (this._worksheet.Names.ContainsKey("_xlnm._FilterDatabase"))
            {
                this._worksheet.Names.Remove("_xlnm._FilterDatabase");
            }
            if (value)
            {
                this.ValidateAutofilterDontCollide();
                ExcelTable? tbl = ExcelTableCollection.GetFromRange(this);
                if (tbl == null)
                {
                    this._worksheet.AutoFilterAddress = this;
                    ExcelNamedRange? result = this._worksheet.Names.AddName("_xlnm._FilterDatabase", this);
                    result.IsNameHidden = true;
                }
                else
                {
                    tbl.ShowFilter = true;
                }
            }
            else
            {
                this._worksheet.AutoFilterAddress = null;
            }
        }
    }

    private void ValidateAutofilterDontCollide()
    {
        foreach (ExcelTable? tbl in this._worksheet.Tables)
        {
            eAddressCollition c = tbl.Address.Collide(this);
            if (c == eAddressCollition.Equal)
            {
                return;   //Autofilter is on a table.
            }

            if (c != eAddressCollition.No)
            {
                throw new InvalidOperationException($"Auto filter collides with table {tbl.Name}");
            }
        }
        foreach (ExcelPivotTable? pt in this._worksheet.PivotTables)
        {
            eAddressCollition c = pt.Address.Collide(this);
            if (c != eAddressCollition.No)
            {
                throw new InvalidOperationException($"Auto filter collides with pivot table {pt.Name}");
            }
        }
    }

    /// <summary>
    /// If the value is in richtext format.
    /// </summary>
    public bool IsRichText
    {
        get
        {
            this.IsRangeValid("richtext");
            return this._worksheet._flags.GetFlagValue(this._fromRow, this._fromCol, CellFlags.RichText);
        }
        set
        {
            this.SetIsRichTextFlag(value);
        }
    }
    /// <summary>
    /// Returns true if the range is a table. If the range partly matches a table range false will be returned.
    /// <seealso cref="IsTable"/>
    /// </summary>
    public bool IsTable
    {
        get
        {
            return ExcelTableCollection.GetFromRange(this) != null;
        }
    }
    /// <summary>
    /// Returns the <see cref="ExcelTable"/> if the range is a table. 
    /// If the range doesn't or partly matches a table range, null is returned.
    /// <seealso cref="IsTable"/>
    /// </summary>
    public ExcelTable GetTable()
    {
        return ExcelTableCollection.GetFromRange(this);
    }
    internal void SetIsRichTextFlag(bool value)
    {
        this._changePropMethod(this, _setIsRichTextDelegate, value);
    }

    /// <summary>
    /// Insert cells into the worksheet and shift the cells to the selected direction.
    /// </summary>
    /// <param name="shift">The direction that the cells will shift.</param>
    public void Insert(eShiftTypeInsert shift)
    {
        if (shift == eShiftTypeInsert.EntireColumn)
        {
            WorksheetRangeInsertHelper.InsertColumn(this._worksheet, this._fromCol, this.Columns, this._fromCol - 1);
        }
        else if (shift == eShiftTypeInsert.EntireRow)
        {
            WorksheetRangeInsertHelper.InsertRow(this._worksheet, this._fromRow, this.Rows, this._fromRow - 1);
        }
        else
        {
            WorksheetRangeInsertHelper.Insert(this, shift, true, false);
        }
    }
    /// <summary>
    /// Delete the range from the worksheet and shift affected cells in the selected direction.
    /// </summary>
    /// <param name="shift">The direction that the cells will shift.</param>
    public void Delete(eShiftTypeDelete shift)
    {
        if (shift == eShiftTypeDelete.EntireColumn || (this._fromRow <= 1 && this._toRow >= ExcelPackage.MaxRows))
        {
            WorksheetRangeDeleteHelper.DeleteColumn(this._worksheet, this._fromCol, this.Columns);
        }
        else if (shift == eShiftTypeDelete.EntireRow || (this._fromCol <= 1 && this._toRow >= ExcelPackage.MaxColumns))
        {
            WorksheetRangeDeleteHelper.DeleteRow(this._worksheet, this._fromRow, this.Rows);
        }
        else
        {
            WorksheetRangeDeleteHelper.Delete(this, shift);
        }
    }

    /// <summary>
    /// Is the range a part of an Arrayformula
    /// </summary>
    public bool IsArrayFormula
    {
        get
        {
            this.IsRangeValid("arrayformulas");
            return this._worksheet._flags.GetFlagValue(this._fromRow, this._fromCol, CellFlags.ArrayFormula);
        }
    }
    /// <summary>
    /// The richtext collection
    /// </summary>
    protected internal ExcelRichTextCollection _rtc = null;
    /// <summary>
    /// The cell value is rich text formatted. 
    /// The RichText-property only apply to the left-top cell of the range.
    /// </summary>
    public ExcelRichTextCollection RichText
    {
        get
        {
            this.IsRangeValid("richtext");

            return this._rtc ??= this._worksheet.GetRichText(this._fromRow, this._fromCol, this);
        }
    }

    /// <summary>
    /// Returns the comment object of the first cell in the range
    /// </summary>
    public ExcelComment Comment
    {
        get
        {
            this.IsRangeValid("comments");
            int i = -1;
            if (this._worksheet.Comments.Count > 0)
            {
                if (this._worksheet._commentsStore.Exists(this._fromRow, this._fromCol, ref i))
                {
                    return this._worksheet._comments._list[i];
                }
            }
            return null;
        }
    }
    /// <summary>
    /// Returns the threaded comment object of the first cell in the range
    /// </summary>
    public ExcelThreadedCommentThread ThreadedComment
    {
        get
        {
            this.IsRangeValid("threaded comments");
            int i = -1;
            if (this._worksheet.ThreadedComments.Count > 0)
            {
                if (this._worksheet._threadedCommentsStore.Exists(this._fromRow, this._fromCol, ref i))
                {
                    return this._worksheet._threadedComments._threads[i];
                }
            }
            return null;
        }
    }
    /// <summary>
    /// WorkSheet object 
    /// </summary>
    public ExcelWorksheet Worksheet
    {
        get
        {
            return this._worksheet;
        }
    }
    /// <summary>
    /// Address including sheet name
    /// </summary>
    public new string FullAddress
    {
        get
        {
            if (this.Addresses == null)
            {
                return GetFullAddress(this._worksheet.Name, this._address);
            }
            else
            {
                string fullAddress = "";
                foreach (ExcelAddressBase? a in this.Addresses)
                {
                    fullAddress += GetFullAddress(this._worksheet.Name, a.Address) + ",";
                }
                return fullAddress.Substring(0, fullAddress.Length - 1);
            }
        }
    }
    /// <summary>
    /// Address including sheetname
    /// </summary>
    public string FullAddressAbsolute
    {
        get
        {
            string wbwsRef = string.IsNullOrEmpty(this._wb) ? this._ws : "[" + this._wb.Replace("'", "''") + "]" + this._ws;
            string fullAddress;
            if (this.Addresses == null)
            {
                fullAddress = GetFullAddress(wbwsRef, GetAddress(this._fromRow, this._fromCol, this._toRow, this._toCol, true));
            }
            else
            {
                fullAddress = "";
                foreach (ExcelAddressBase? a in this.Addresses)
                {
                    if (fullAddress != "")
                    {
                        fullAddress += ",";
                    }

                    if (a.Address == "#REF!")
                    {
                        fullAddress += GetFullAddress(wbwsRef, "#REF!");
                    }
                    else
                    {
                        fullAddress += GetFullAddress(wbwsRef, GetAddress(a.Start.Row, a.Start.Column, a.End.Row, a.End.Column, true));
                    }
                }
            }
            return fullAddress;
        }
    }
    #endregion
    #region Private Methods
    /// <summary>
    /// Set the value without altering the richtext property
    /// </summary>
    /// <param name="value">the value</param>
    internal void SetValueRichText(object value)
    {
        if (this._fromRow == 1 && this._fromCol == 1 && this._toRow == ExcelPackage.MaxRows && this._toCol == ExcelPackage.MaxColumns)  //Full sheet (ex ws.Cells.Value=0). Set value for A1 only to avoid hanging 
        {
            this.SetValueInner(value, 1, 1);
        }
        else
        {
            this.SetValueInner(value, this._fromRow, this._fromCol);
        }
    }

    private void SetValueInner(object value, int row, int col)
    {
        this._worksheet.SetValue(row, col, value);
        this._worksheet._formulas.SetValue(row, col, "");
    }
    internal void SetSharedFormulaID(int id, int prevId)
    {
        for (int col = this._fromCol; col <= this._toCol; col++)
        {
            for (int row = this._fromRow; row <= this._toRow; row++)
            {
                int? f = this._worksheet._formulas.GetValue(row, col) as int?;
                if (f.HasValue && f.Value == prevId)
                {
                    this._worksheet._formulas.SetValue(row, col, id);
                }
            }
        }
    }
    private void CheckAndSplitSharedFormula(ExcelAddressBase address)
    {
        for (int col = address._fromCol; col <= address._toCol; col++)
        {
            for (int row = address._fromRow; row <= address._toRow; row++)
            {
                object? f = this._worksheet._formulas.GetValue(row, col);
                if (f is int && (int)f >= 0)
                {
                    this.SplitFormulas(address);
                    return;
                }
            }
        }
    }

    private void SplitFormulas(ExcelAddressBase address)
    {
        List<int> formulas = new List<int>();
        for (int col = address._fromCol; col <= address._toCol; col++)
        {
            for (int row = address._fromRow; row <= address._toRow; row++)
            {
                object? f = this._worksheet._formulas.GetValue(row, col);
                if (f is int)
                {
                    int id = (int)f;
                    if (id >= 0 && !formulas.Contains(id))
                    {
                        if (this._worksheet._sharedFormulas[id].FormulaType==ExcelWorksheet.FormulaType.Array && this.Collide(this._worksheet.Cells[this._worksheet._sharedFormulas[id].Address]) == eAddressCollition.Partly) // If the formula is an array formula and its on the inside the overwriting range throw an exception
                        {
                            throw (new InvalidOperationException("Cannot overwrite a part of an array-formula"));
                        }
                        formulas.Add(id);
                    }
                }
            }
        }

        foreach (int ix in formulas)
        {
            this.SplitFormula(address, ix);
        }

        ////Clear any formula references inside the refered range
        //_worksheet._formulas.Clear(address._fromRow, address._toRow, address._toRow - address._fromRow + 1, address._toCol - address.column + 1);
    }

    private void SplitFormula(ExcelAddressBase address, int ix)
    {
        ExcelWorksheet.Formulas? f = this._worksheet._sharedFormulas[ix];
        ExcelRange? fRange = this._worksheet.Cells[f.Address];
        eAddressCollition collide = address.Collide(fRange);

        //The formula is inside the currenct range, remove it
        if (collide == eAddressCollition.Equal || collide == eAddressCollition.Inside)
        {
            this._worksheet._sharedFormulas.Remove(ix);
            return;
            //fRange.SetSharedFormulaID(int.MinValue); 
        }
        eAddressCollition firstCellCollide = address.Collide(new ExcelAddressBase(fRange._fromRow, fRange._fromCol, fRange._fromRow, fRange._fromCol));
        if (collide == eAddressCollition.Partly && (firstCellCollide == eAddressCollition.Inside || firstCellCollide == eAddressCollition.Equal)) //Do we need to split? Only if the functions first row is inside the new range.
        {
            //The formula partly collides with the current range
            bool fIsSet = false;
            string formulaR1C1 = fRange.FormulaR1C1;
            //Top Range
            if (fRange._fromRow < this._fromRow)
            {
                f.Address = GetAddress(fRange._fromRow, fRange._fromCol, this._fromRow - 1, fRange._toCol);
                fIsSet = true;
            }
            int pIx = f.Index;
            //Left Range
            if (fRange._fromCol < address._fromCol)
            {
                if (fIsSet)
                {
                    f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default);
                    f.Index = this._worksheet.GetMaxShareFunctionIndex(false);
                    f.StartCol = fRange._fromCol;
                    f.FormulaType = ExcelWorksheet.FormulaType.Shared;
                    this._worksheet._sharedFormulas.Add(f.Index, f);
                }
                else
                {
                    fIsSet = true;
                }
                if (fRange._fromRow < address._fromRow)
                {
                    f.StartRow = address._fromRow;
                }
                else
                {
                    f.StartRow = fRange._fromRow;
                }
                if (fRange._toRow < address._toRow)
                {
                    f.Address = GetAddress(f.StartRow, f.StartCol,
                                           fRange._toRow, address._fromCol - 1);
                }
                else
                {
                    f.Address = GetAddress(f.StartRow, f.StartCol,
                                           address._toRow, address._fromCol - 1);
                }
                f.Formula = TranslateFromR1C1(formulaR1C1, f.StartRow, f.StartCol);
                this._worksheet.Cells[f.Address].SetSharedFormulaID(f.Index, pIx);
            }
            //Right Range
            if (fRange._toCol > address._toCol)
            {
                if (fIsSet)
                {
                    f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default);
                    f.Index = this._worksheet.GetMaxShareFunctionIndex(false);
                    f.FormulaType = ExcelWorksheet.FormulaType.Shared;
                    this._worksheet._sharedFormulas.Add(f.Index, f);
                }
                else
                {
                    fIsSet = true;
                }
                f.StartCol = address._toCol + 1;
                if (address._fromRow < fRange._fromRow)
                {
                    f.StartRow = fRange._fromRow;
                }
                else
                {
                    f.StartRow = address._fromRow;
                }

                if (fRange._toRow < address._toRow)
                {
                    f.Address = GetAddress(f.StartRow, f.StartCol,
                                           fRange._toRow, fRange._toCol);
                }
                else
                {
                    f.Address = GetAddress(f.StartRow, f.StartCol,
                                           address._toRow, fRange._toCol);
                }
                f.Formula = TranslateFromR1C1(formulaR1C1, f.StartRow, f.StartCol);
                this._worksheet.Cells[f.Address].SetSharedFormulaID(f.Index, pIx);
            }
            //Bottom Range
            if (fRange._toRow > address._toRow)
            {
                if (fIsSet)
                {
                    f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default);
                    f.Index = this._worksheet.GetMaxShareFunctionIndex(false);
                    f.FormulaType = ExcelWorksheet.FormulaType.Shared;
                    this._worksheet._sharedFormulas.Add(f.Index, f);
                }

                f.StartCol = fRange._fromCol;
                f.StartRow = address._toRow + 1;

                f.Formula = TranslateFromR1C1(formulaR1C1, f.StartRow, f.StartCol);

                f.Address = GetAddress(f.StartRow, f.StartCol,
                                       fRange._toRow, fRange._toCol);

                this._worksheet.Cells[f.Address].SetSharedFormulaID(f.Index, pIx);

            }
        }
    }

    /// <summary>
    /// Removes all formulas within the range, but keeps the calculated values.
    /// </summary>
    public void ClearFormulas()
    {
        CellStoreEnumerator<object>? formulaCells = new CellStoreEnumerator<object>(this.Worksheet._formulas, this.Start.Row, this.Start.Column, this.End.Row, this.End.Column);
        while (formulaCells.Next())
        {
            formulaCells.Value = null;
        }

        this.ClearTableFormulas();
    }

    private void ClearTableFormulas()
    {
        //Clear any calculated formulas in tables.
        foreach (ExcelTable? table in this.Worksheet.Tables)
        {
            if (table.Address.Collide(this) != eAddressCollition.No)
            {
                foreach (ExcelTableColumn? col in table.Columns)
                {
                    if (string.IsNullOrEmpty(col.CalculatedColumnFormula) == false)
                    {
                        ExcelAddressBase? adr = table.Address;
                        int fromRow = table.ShowHeader ? adr._fromRow + 1 : adr._fromRow;
                        int toRow = table.ShowTotal ? adr._toRow - 1 : adr._toRow;
                        int colNum = adr._fromCol + col.Position;
                        CellStoreEnumerator<object>? formulaCells = new CellStoreEnumerator<object>(this.Worksheet._formulas, fromRow, colNum, toRow, colNum);
                        bool hasValue = false;
                        while (formulaCells.Next())
                        {
                            if (formulaCells.Value != null && 
                                formulaCells.Value.ToString().Equals(col.CalculatedColumnFormula, StringComparison.OrdinalIgnoreCase))
                            {
                                hasValue = true;
                                break;
                            }
                        }
                        if (hasValue == false)
                        {
                            col.RemoveFormulaNode();
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// Removes all values of cells with formulas, but keeps the formulas.
    /// </summary>
    public void ClearFormulaValues()
    {
        CellStoreEnumerator<object>? formulaCell = new CellStoreEnumerator<object>(this.Worksheet._formulas, this.Start.Row, this.Start.Column, this.End.Row, this.End.Column);
        while (formulaCell.Next())
        {
            ExcelValue val = this.Worksheet._values.GetValue(formulaCell.Row, formulaCell.Column);
            val._value = null;
            this.Worksheet._values.SetValue(formulaCell.Row, formulaCell.Column, val);
        }
    }

    private static object ConvertData(ExcelTextFormat Format, string v, int col, bool isText)
    {
        if (isText && (Format.DataTypes == null || Format.DataTypes.Length < col))
        {
            return string.IsNullOrEmpty(v) ? null : v;
        }

        double d;
        DateTime dt;
        if (Format.DataTypes == null || Format.DataTypes.Length <= col || Format.DataTypes[col] == eDataTypes.Unknown)
        {
            string v2 = v.EndsWith("%") ? v.Substring(0, v.Length - 1) : v;
            if (double.TryParse(v2, NumberStyles.Any, Format.Culture, out d))
            {
                if (v2 == v)
                {
                    return d;
                }
                else
                {
                    return d / 100;
                }
            }
            if (DateTime.TryParse(v, Format.Culture, DateTimeStyles.None, out dt))
            {
                return dt;
            }
            else
            {
                return string.IsNullOrEmpty(v) ? null : v;
            }
        }
        else
        {
            switch (Format.DataTypes[col])
            {
                case eDataTypes.Number:
                    if (double.TryParse(v, NumberStyles.Any, Format.Culture, out d))
                    {
                        return d;
                    }
                    else
                    {
                        return v;
                    }
                case eDataTypes.DateTime:
                    if (DateTime.TryParse(v, Format.Culture, DateTimeStyles.None, out dt))
                    {
                        return dt;
                    }
                    else
                    {
                        return v;
                    }
                case eDataTypes.Percent:
                    string v2 = v.EndsWith("%") ? v.Substring(0, v.Length - 1) : v;
                    if (double.TryParse(v2, NumberStyles.Any, Format.Culture, out d))
                    {
                        return d / 100;
                    }
                    else
                    {
                        return v;
                    }
                case eDataTypes.String:
                    return v;
                default:
                    return string.IsNullOrEmpty(v) ? null : v;

            }
        }
    }
    #endregion
    #region Public Methods
    #region ConditionalFormatting
    /// <summary>
    /// Conditional Formatting for this range.
    /// </summary>
    public IRangeConditionalFormatting ConditionalFormatting
    {
        get
        {
            return new RangeConditionalFormatting(this._worksheet, new ExcelAddress(this.Address));
        }
    }
    #endregion
    #region DataValidation
    /// <summary>
    /// Data validation for this range.
    /// </summary>
    public IRangeDataValidation DataValidation
    {
        get
        {
            return new RangeDataValidation(this._worksheet, this.Address);
        }
    }
    #endregion
    #region GetValue

    /// <summary>
    ///     Convert cell value to desired type, including nullable structs.
    ///     When converting blank string to nullable struct (e.g. ' ' to int?) null is returned.
    ///     When attempted conversion fails exception is passed through.
    /// </summary>
    /// <typeparam name="T">
    ///     The type to convert to.
    /// </typeparam>
    /// <returns>
    ///     The <see cref="Value"/> converted to <typeparamref name="T"/>.
    /// </returns>
    /// <remarks>
    ///     If  <see cref="Value"/> is string, parsing is performed for output types of DateTime and TimeSpan, which if fails throws <see cref="FormatException"/>.
    ///     Another special case for output types of DateTime and TimeSpan is when input is double, in which case <see cref="DateTime.FromOADate"/>
    ///     is used for conversion. This special case does not work through other types convertible to double (e.g. integer or string with number).
    ///     In all other cases 'direct' conversion <see cref="Convert.ChangeType(object, Type)"/> is performed.
    /// </remarks>
    /// <exception cref="FormatException">
    ///      <see cref="Value"/> is string and its format is invalid for conversion (parsing fails)
    /// </exception>
    /// <exception cref="InvalidCastException">
    ///      <see cref="Value"/> is not string and direct conversion fails
    /// </exception>
    public T GetValue<T>()
    {
        return ConvertUtil.GetTypedCellValue<T>(this.Value);
    }
    #endregion
    /// <summary>
    /// Get a range with an offset from the top left cell.
    /// The new range has the same dimensions as the current range
    /// </summary>
    /// <param name="RowOffset">Row Offset</param>
    /// <param name="ColumnOffset">Column Offset</param>
    /// <returns></returns>
    public ExcelRangeBase Offset(int RowOffset, int ColumnOffset)
    {
        if (this._fromRow + RowOffset < 1 || this._fromCol + ColumnOffset < 1 || this._fromRow + RowOffset > ExcelPackage.MaxRows || this._fromCol + ColumnOffset > ExcelPackage.MaxColumns)
        {
            throw (new ArgumentOutOfRangeException("Offset value out of range"));
        }
        string address = GetAddress(this._fromRow + RowOffset, this._fromCol + ColumnOffset, this._toRow + RowOffset, this._toCol + ColumnOffset);
        return new ExcelRangeBase(this._worksheet, address);
    }
    /// <summary>
    /// Get a range with an offset from the top left cell.
    /// </summary>
    /// <param name="RowOffset">Row Offset</param>
    /// <param name="ColumnOffset">Column Offset</param>
    /// <param name="NumberOfRows">Number of rows. Minimum 1</param>
    /// <param name="NumberOfColumns">Number of colums. Minimum 1</param>
    /// <returns></returns>
    public ExcelRangeBase Offset(int RowOffset, int ColumnOffset, int NumberOfRows, int NumberOfColumns)
    {
        if (NumberOfRows < 1 || NumberOfColumns < 1)
        {
            throw (new Exception("Number of rows/columns must be greater than 0"));
        }
        NumberOfRows--;
        NumberOfColumns--;
        if (this._fromRow + RowOffset < 1 || this._fromCol + ColumnOffset < 1 || this._fromRow + RowOffset > ExcelPackage.MaxRows || this._fromCol + ColumnOffset > ExcelPackage.MaxColumns || this._fromRow + RowOffset + NumberOfRows < 1 || this._fromCol + ColumnOffset + NumberOfColumns < 1 || this._fromRow + RowOffset + NumberOfRows > ExcelPackage.MaxRows || this._fromCol + ColumnOffset + NumberOfColumns > ExcelPackage.MaxColumns)
        {
            throw (new ArgumentOutOfRangeException("Offset value out of range"));
        }
        string address = GetAddress(this._fromRow + RowOffset, this._fromCol + ColumnOffset, this._fromRow + RowOffset + NumberOfRows, this._fromCol + ColumnOffset + NumberOfColumns);
        return new ExcelRangeBase(this._worksheet, address);
    }
    /// <summary>
    /// Adds a new comment for the range.
    /// If this range contains more than one cell, the top left comment is returned by the method.
    /// </summary>
    /// <param name="Text">The text for the comment</param>
    /// <param name="Author">The author for the comment. If this property is null or blank EPPlus will set it to the identity of the ClaimsPrincipal if available otherwise to "Anonymous"</param>
    /// <returns>A reference comment of the top left cell</returns>
    public ExcelComment AddComment(string Text, string Author = null)
    {
        //Check if any comments exists in the range and throw an exception
        this._changePropMethod(this, _setExistsCommentDelegate, null);
        //Create the comments
        this._changePropMethod(this, _setCommentDelegate, new string[] { Text, Author });

        return this._worksheet.Comments[new ExcelCellAddress(this._fromRow, this._fromCol)];
    }
    /// <summary>
    /// Adds a new threaded comment for the range.
    /// If this range contains more than one cell, the top left comment is returned by the method.
    /// </summary>
    /// <returns>A reference comment of the top left cell</returns>
    public ExcelThreadedCommentThread AddThreadedComment()
    {
        //Check if any comments exists in the range and throw an exception
        this._changePropMethod(this, _setExistsThreadedCommentDelegate, null);
        //Create the comments
        this._changePropMethod(this, _setThreadedCommentDelegate, new string[0]);

        return this._worksheet.ThreadedComments[new ExcelCellAddress(this._fromRow, this._fromCol)];
    }

    /// <summary>
    /// Copies the range of cells to another range. 
    /// </summary>
    /// <param name="Destination">The top-left cell where the range will be copied.</param>
    public void Copy(ExcelRangeBase Destination)
    {
        RangeCopyHelper? helper = new RangeCopyHelper(this, Destination, 0);
        helper.Copy();
    }

    /// <summary>
    /// Copies the range of cells to an other range
    /// </summary>
    /// <param name="Destination">The start cell where the range will be copied.</param>
    /// <param name="excelRangeCopyOptionFlags">Cell properties that will not be copied.</param>
    public void Copy(ExcelRangeBase Destination, ExcelRangeCopyOptionFlags? excelRangeCopyOptionFlags)
    {
        RangeCopyHelper? helper = new RangeCopyHelper(this, Destination, excelRangeCopyOptionFlags ?? 0);
        helper.Copy();
    }
    /// <summary>
    /// Copies the range of cells to an other range
    /// </summary>
    /// <param name="Destination">The start cell where the range will be copied.</param>
    /// <param name="excelRangeCopyOptionFlags">Cell properties that will not be copied.</param>
    public void Copy(ExcelRangeBase Destination, params ExcelRangeCopyOptionFlags[] excelRangeCopyOptionFlags)
    {
        ExcelRangeCopyOptionFlags flags=0;
        foreach (ExcelRangeCopyOptionFlags c in excelRangeCopyOptionFlags)
        {
            flags |= c;
        }
        RangeCopyHelper? helper = new RangeCopyHelper(this, Destination, flags);
        helper.Copy();
    }
    /// <summary>
    /// Copy the styles from the source range to the destination range.
    /// If the destination range is larger than the source range, the styles of the column to the right and the row at the bottom will be expanded to the destination.
    /// </summary>
    /// <param name="Destination">The destination range</param>
    public void CopyStyles(ExcelRangeBase Destination)
    {
        RangeCopyStylesHelper? helper = new RangeCopyStylesHelper(this, Destination);
        helper.CopyStyles();
    }
    /// <summary>
    /// Clear all cells
    /// </summary>
    public void Clear()
    {
        this.DeleteMe(this, false);
    }
    /// <summary>
    /// Creates an array-formula.
    /// </summary>
    /// <param name="ArrayFormula">The formula</param>
    public void CreateArrayFormula(string ArrayFormula)
    {
        if (this.Addresses != null)
        {
            throw (new Exception("An array formula cannot have more than one address"));
        }
        Set_SharedFormula(this, ArrayFormula, this, true);
    }
    internal void DeleteMe(ExcelAddressBase Range, bool shift, bool clearValues = true, bool clearFormulas = true, bool clearFlags = true, bool clearMergedCells = true, bool clearHyperLinks = true, bool clearComments = true, bool clearThreadedComments=true, bool clearStyles=true)
    {

        //First find the start cell
        FormulaDataTableValidation.HasPartlyFormulaDataTable(this._worksheet, Range, false, "Can't clear a part of a data table function");
            
        int fromRow, fromCol;
        ExcelAddressBase? d = this.Worksheet.Dimension;
        if (d != null && Range._fromRow <= d._fromRow && Range._toRow >= d._toRow) //EntireRow?
        {
            fromRow = d._fromRow;
        }
        else
        {
            fromRow = Range._fromRow;
        }
        if (d != null && Range._fromCol <= d._fromCol && Range._toCol >= d._toCol) //EntireRow?
        {
            fromCol = d._fromCol;
        }
        else
        {
            fromCol = Range._fromCol;
        }

        int rows = Range._toRow - fromRow + 1;
        int cols = Range._toCol - fromCol + 1;


        if (clearMergedCells)
        {
            this._worksheet.MergedCells.Clear(Range);
        }

        if (clearValues && clearStyles)
        {
            this._worksheet._values.Delete(fromRow, fromCol, rows, cols, shift);
        }
        else if(clearValues)
        {
            ClearValue(this._worksheet._values, true, fromRow, fromCol, rows, cols);
        }
        else if(clearStyles)
        {
            ClearValue(this._worksheet._values, false, fromRow, fromCol, rows, cols);
        }

        if (clearFormulas)
        {
            this._worksheet._formulas.Delete(fromRow, fromCol, rows, cols, shift);
        }

        if (clearFlags)
        {
            this._worksheet._flags.Delete(fromRow, fromCol, rows, cols, shift);
            this._worksheet._metadataStore.Delete(fromRow, fromCol, rows, cols, shift);
        }
        if (clearHyperLinks)
        {
            this._worksheet._hyperLinks.Delete(fromRow, fromCol, rows, cols, shift);
        }
        if (clearComments)
        {
            this.DeleteComments(Range);
        }
        if (clearThreadedComments)
        {
            this.DeleteThreadedComments(Range);
        }

        //Clear multi addresses as well
        if (Range.Addresses != null)
        {
            foreach (ExcelAddressBase? sub in Range.Addresses)
            {
                this.DeleteMe(sub, shift, clearValues, clearFormulas, clearFlags, clearMergedCells, clearHyperLinks, clearComments, clearThreadedComments, clearStyles);
            }
        }
    }

    /// <summary>
    /// Clears either value or style for a range from the cellstore.
    /// </summary>
    /// <param name="values"></param>
    /// <param name="clearValue"></param>
    /// <param name="fromRow"></param>
    /// <param name="fromCol"></param>
    /// <param name="rows"></param>
    /// <param name="cols"></param>
    private static void ClearValue(CellStoreValue values, bool clearValue, int fromRow, int fromCol, int rows, int cols)
    {
        int toRow = fromRow + rows - 1;
        int toCol = fromCol + cols - 1;
        CellStoreEnumerator<ExcelValue>? cse = new CellStoreEnumerator<ExcelValue>(values, fromRow, fromCol, toRow, toCol);
        while (cse.Next())
        {
            if(clearValue)
            {
                cse.Value = new ExcelValue() { _value = null, _styleId = cse.Value._styleId };
            }
            else
            {
                cse.Value = new ExcelValue() { _value = cse.Value, _styleId = 0 };
            }
        }
    }

    private void DeleteComments(ExcelAddressBase Range)
    {
        List<int>? deleted = new List<int>();
        CellStoreEnumerator<int>? cse = new CellStoreEnumerator<int>(this._worksheet._commentsStore, Range._fromRow, Range._fromCol, Range._toRow, Range._toCol);
        while (cse.Next())
        {
            if (this._worksheet._threadedCommentsStore.Exists(cse.Row, cse.Column) == false) //Threaded comments keep a comment for backward compatibility that needs to be keept.
            {
                deleted.Add(cse.Value);
            }
        }
        foreach (int i in deleted)
        {
            this._worksheet.Comments.Remove(this._worksheet.Comments._list[i]);
        }
    }
    private void DeleteThreadedComments(ExcelAddressBase Range)
    {
        List<int>? deleted = new List<int>();
        CellStoreEnumerator<int>? cse = new CellStoreEnumerator<int>(this._worksheet._threadedCommentsStore, Range._fromRow, Range._fromCol, Range._toRow, Range._toCol);
        while (cse.Next())
        {
            deleted.Add(cse.Value);
        }
        foreach (int i in deleted)
        {
            this._worksheet.ThreadedComments.Remove(this._worksheet.ThreadedComments._threads[i]);
        }
    }

    #endregion
    #region IDisposable Members
    /// <summary>
    /// Disposes the object
    /// </summary>
    public void Dispose()
    {
        //_worksheet = null;            
    }

    #endregion
    #region "Enumerator"
    CellStoreEnumerator<ExcelValue> cellEnum;
    /// <summary>
    /// Gets the enumerator for the collection
    /// </summary>
    /// <returns>The enumerator</returns>
    public IEnumerator<ExcelRangeBase> GetEnumerator()
    {
        this.Reset();
        return this;
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        this.Reset();
        return this;
    }

    /// <summary>
    /// The current range when enumerating
    /// </summary>
    public ExcelRangeBase Current
    {
        get
        {
            if (this.cellEnum == null)
            {
                return null;
            }
            return new ExcelRangeBase(this._worksheet, GetAddress(this.cellEnum.Row, this.cellEnum.Column));
        }
    }

    /// <summary>
    /// The current range when enumerating
    /// </summary>
    object IEnumerator.Current
    {
        get
        {
            if (this.cellEnum == null)
            {
                return null;
            }
            return ((object)(new ExcelRangeBase(this._worksheet, GetAddress(this.cellEnum.Row, this.cellEnum.Column))));
        }
    }

    //public object FormatedText { get; private set; }

    int _enumAddressIx = 0;
    /// <summary>
    /// Iterate to the next cell
    /// </summary>
    /// <returns>False if no more cells exists</returns>
    public bool MoveNext()
    {
        if (this.cellEnum == null)
        {
            this.Reset();
        }

        if (this.cellEnum.Next())
        {
            return true;
        }
        else if (this._addresses != null)
        {
            this._enumAddressIx++;
            if (this._enumAddressIx < this._addresses.Count)
            {
                this.cellEnum = new CellStoreEnumerator<ExcelValue>(this._worksheet._values,
                                                                    this._addresses[this._enumAddressIx]._fromRow,
                                                                    this._addresses[this._enumAddressIx]._fromCol,
                                                                    this._addresses[this._enumAddressIx]._toRow,
                                                                    this._addresses[this._enumAddressIx]._toCol);
                return this.MoveNext();
            }
            else
            {
                return false;
            }
        }
        return false;
    }
    /// <summary>
    /// Reset the enumerator
    /// </summary>
    public void Reset()
    {
        this._enumAddressIx = 0;
        this.cellEnum = new CellStoreEnumerator<ExcelValue>(this._worksheet._values, this._fromRow, this._fromCol, this._toRow, this._toCol);
    }
    #endregion

    /// <summary>
    /// Sort the range by value of the first column, Ascending.
    /// </summary>
    public void Sort()
    {
        this.SortInternal(new int[] { 0 }, new bool[] { false }, null, null, CompareOptions.None, null);
    }
    /// <summary>
    /// Sort the range by value of the supplied column, Ascending.
    /// <param name="column">The column to sort by within the range. Zerobased</param>
    /// <param name="descending">Descending if true, otherwise Ascending. Default Ascending. Zerobased</param>
    /// </summary>
    public void Sort(int column, bool descending = false)
    {
        this.SortInternal(new int[] { column }, new bool[] { descending }, null, null, CompareOptions.None, null);
    }
    /// <summary>
    /// Sort the range by value
    /// </summary>
    /// <param name="columns">The column(s) to sort by within the range. Zerobased</param>
    /// <param name="descending">Descending if true, otherwise Ascending. Default Ascending. Zerobased</param>
    /// <param name="culture">The CultureInfo used to compare values. A null value means CurrentCulture</param>
    /// <param name="compareOptions">String compare option</param>
    public void Sort(int[] columns, bool[] descending = null, CultureInfo culture = null, CompareOptions compareOptions = CompareOptions.None)
    {
        this.SortInternal(columns, descending, null, culture, compareOptions, null);
    }

    /// <summary>
    /// Sort the range by value
    /// </summary>
    /// <param name="columns">The column(s) to sort by within the range. Zerobased</param>
    /// <param name="descending">Descending if true, otherwise Ascending. Default Ascending. Zerobased</param>
    /// <param name="customLists">A Dictionary containing custom lists indexed by column</param>
    /// <param name="culture">The CultureInfo used to compare values. A null value means CurrentCulture</param>
    /// <param name="compareOptions">String compare option</param>
    /// <param name="table"><see cref="ExcelTable"/> to be sorted</param>
    /// <param name="leftToRight">Indicates if the range should be sorted left to right (by column) instead of top-down (by row)</param>
    internal void SortInternal(
        int[] columns,
        bool[] descending = null,
        Dictionary<int, string[]> customLists = null,
        CultureInfo culture = null,
        CompareOptions compareOptions = CompareOptions.None,
        ExcelTable table = null,
        bool leftToRight = false)
    {
        if (leftToRight)
        {
            this._worksheet._rangeSorter.SortLeftToRight(this, columns, ref descending, culture, compareOptions, customLists);
        }
        else
        {
            this._worksheet._rangeSorter.Sort(this, columns, ref descending, culture, compareOptions, customLists);
        }

        if (table != null)
        {
            table.SetTableSortState(columns, descending, compareOptions, customLists);
        }
        else
        {
            this._worksheet._rangeSorter.SetWorksheetSortState(this, columns, descending, compareOptions, leftToRight, customLists);
        }
    }

    /// <summary>
    /// Sort the range by value
    /// </summary>
    /// <param name="options">An instance of <see cref="RangeSortOptions"/> where sort parameters can be set</param>
    internal void SortInternal(SortOptionsBase options)
    {
        if (options.ColumnIndexes.Count > 0)
        {
            this.SortInternal(options.ColumnIndexes.ToArray(), options.Descending.ToArray(), options.CustomLists, options.Culture, options.CompareOptions, null, options.LeftToRight);
        }
        else
        {
            this.Sort(new int[] { 0 }, new bool[] { false }, options.Culture, options.CompareOptions);
        }
    }

    internal void Sort(SortOptionsBase options, ExcelTable table)
    {
        this.SortInternal(options.ColumnIndexes.ToArray(), options.Descending.ToArray(), options.CustomLists, options.Culture, options.CompareOptions, table);
    }

    /// <summary>
    /// Sort the range by value. Supports top-down and left to right sort.
    /// </summary>
    /// <param name="configuration">An action of <see cref="RangeSortOptions"/> where sort parameters can be set.</param>
    /// <example> 
    /// <code>
    /// // 1. Sort rows (top-down)
    /// 
    /// // The Column function takes the zero based column index in the range
    /// worksheet.Cells["A1:D15"].Sort(x => x.SortBy.Column(0).ThenSortBy.Column(1, eSortOrder.Descending));
    /// 
    /// // 2. Sort columns(left to right)
    /// // The Row function takes the zero based row index in the range
    /// worksheet.Cells["A1:D15"].Sort(x => x.SortLeftToRightBy.Row(0));
    /// 
    /// // 3. Sort using a custom list
    /// worksheet.Cells["A1:D15"].Sort(x => x.SortBy.Column(0).UsingCustomList("S", "M", "L", "XL"));
    /// worksheet.Cells["A1:D15"].Sort(x => x.SortLeftToRightBy.Row(0).UsingCustomList("S", "M", "L", "XL"));
    /// </code>
    /// </example>
    public void Sort(Action<RangeSortOptions> configuration)
    {
        RangeSortOptions? options = new RangeSortOptions();
        configuration(options);
        this.SortInternal(options);
    }

    /// <summary>
    /// Sort the range by value. Use RangeSortOptions.Create() to create an instance of the sort options, then
    /// use the <see cref="RangeSortOptions.SortBy"/> or <see cref="RangeSortOptions.SortLeftToRightBy"/> properties to build up your sort parameters.
    /// </summary>
    /// <param name="options"><see cref="RangeSortOptions">Options</see> for the sort</param>
    /// <example> 
    /// <code>
    /// var options = RangeSortOptions.Create();
    /// var builder = options.SortBy.Column(0);
    /// builder.ThenSortBy.Column(2).UsingCustomList("S", "M", "L", "XL");
    /// builder.ThenSortBy.Column(3);
    /// worksheet.Cells["A1:D15"].Sort(options);
    /// </code>
    /// </example>
    public void Sort(RangeSortOptions options)
    {
        this.SortInternal(options);
    }

    private static void SortSetValue(List<ExcelValue> list, int index, object value)
    {
        ExcelValue v = (ExcelValue)value;
        list[index] = new ExcelValue { _value = v._value, _styleId = v._styleId };
    }
    /// <summary>
    /// If the range is a name or a table, return the name.
    /// </summary>
    /// <returns></returns>
    internal string GetName()
    {
        if (this is ExcelNamedRange n)
        {
            return n.Name;
        }
        else
        {
            ExcelTable? t = ExcelTableCollection.GetFromRange(this);
            if (t != null)
            {
                return t.Name;
            }
        }
        return null;
    }
    ExcelRangeColumn _entireColumn = null;
    /// <summary>
    /// A reference to the column properties for column(s= referenced by this range.
    /// If multiple ranges are addressed (e.g a1:a2,c1:c3), only the first address is used.
    /// </summary>
    public ExcelRangeColumn EntireColumn
    {
        get
        {
            if (this._entireColumn == null || this._entireColumn._fromCol != this._fromCol || this._entireColumn._toCol != this._toCol)
            {
                this._entireColumn = new ExcelRangeColumn(this._worksheet, this._fromCol, this._toCol);
            }
            return this._entireColumn;
        }
    }
    ExcelRangeRow _entireRow = null;
    /// <summary>
    /// A reference to the row properties for row(s) referenced by this range.
    /// If multiple ranges are addressed (e.g a1:a2,c1:c3), only the first address is used.
    /// </summary>
    public ExcelRangeRow EntireRow
    {
        get
        {
            if (this._entireRow == null || this._entireRow._fromRow != this._fromRow || this._entireRow._toRow != this._toRow)
            {
                this._entireRow = new ExcelRangeRow(this._worksheet, this._fromRow, this._toRow);
            }
            return this._entireRow;
        }
    }
    /// <summary>
    /// Gets the typed value of a cell 
    /// </summary>
    /// <typeparam name="T">The returned type</typeparam>
    /// <returns>The value of the cell</returns>
    public T GetCellValue<T>()
    {
        return this.GetCellValue<T>(0, 0);
    }
    /// <summary>
    /// Gets the value of a cell using an offset from the top-left cell in the range.
    /// </summary>
    /// <typeparam name="T">The returned type</typeparam>
    /// <param name="columnOffset">Column offset from the top-left cell in the range</param>
    public T GetCellValue<T>(int columnOffset)
    {
        return this.GetCellValue<T>(0, columnOffset);
    }
    /// <summary>
    /// Gets the value of a cell using an offset from the top-left cell in the range.
    /// </summary>
    /// <typeparam name="T">The returned type</typeparam>
    /// <param name="rowOffset">Row offset from the top-left cell in the range</param>
    /// <param name="columnOffset">Column offset from the top-left cell in the range</param>
    public T GetCellValue<T>(int rowOffset, int columnOffset)
    {
        if (this.IsName)
        {
            ExcelNamedRange n;
            if (this._worksheet == null)
            {
                n = this._workbook._names[this._address];
            }
            else
            {

                n = this._worksheet.Names[this._address];
            }
            ExcelAddressBase? a = new ExcelAddressBase(n.Address);
            if (a._fromRow > 0 && a._fromCol > 0)
            {
                return this._worksheet.GetValue<T>(this._fromRow + rowOffset, this._fromCol + columnOffset);
            }
            else
            {
                return default(T);
            }
        }
        else
        {
            return this._worksheet.GetValue<T>(this._fromRow + rowOffset, this._fromCol + columnOffset);
        }
    } 
    /// <summary>
    /// Sets the value of a cell using an offset from the top-left cell in the range.
    /// </summary>
    /// <param name="rowOffset">Row offset from the top-left cell in the range</param>
    /// <param name="columnOffset">Column offset from the top-left cell in the range</param>
    /// <param name="value">The value to set.</param>
    public void SetCellValue(int rowOffset, int columnOffset, object value)
    {
        if (this.IsName)
        {
            ExcelNamedRange n;
            if (this._worksheet == null)
            {
                n= this._workbook._names[this._address];
            }
            else
            {
                    
                n= this._worksheet.Names[this._address];
            }                
            ExcelAddressBase? a = new ExcelAddressBase(n.Address);
            if (a._fromRow>0 && a._fromCol>0)
            {
                this._worksheet.SetValue(a._fromRow + rowOffset, a._fromCol + columnOffset, value);
            }
            else
            {
                throw new InvalidOperationException($"Can't set value on name {n.Name} referencing {n.Address}. Offset is not possible.");
            }
        }
        else
        {
            this._worksheet.SetValue(this._fromRow + rowOffset, this._fromCol + columnOffset, value);
        }
    }
}