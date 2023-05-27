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
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table;

/// <summary>
/// A table column
/// </summary>
public class ExcelTableColumn : ExcelTableDxfBase
{
    internal ExcelTable _tbl;
    internal ExcelTableColumn(XmlNamespaceManager ns, XmlNode topNode, ExcelTable tbl, int pos) :
        base(ns, topNode)
    {
        this._tbl = tbl;
        this.InitDxf(tbl.WorkSheet.Workbook.Styles, null, this);
        this.Position = pos;
    }
    /// <summary>
    /// The column id
    /// </summary>
    public int Id 
    {
        get
        {
            return this.GetXmlNodeInt("@id");
        }
        set
        {
            this.SetXmlNodeString("@id", value.ToString());
        }
    }
    /// <summary>
    /// The position of the column
    /// </summary>
    public int Position
    {
        get;
        internal set;
    }
    /// <summary>
    /// The name of the column
    /// </summary>
    public string Name
    {
        get
        {
            string? n= this.GetXmlNodeString("@name");
            if (string.IsNullOrEmpty(n))
            {
                if (this._tbl.ShowHeader)
                {
                    n = ConvertUtil.ExcelDecodeString(this._tbl.WorkSheet.GetValue<string>(this._tbl.Address._fromRow, this._tbl.Address._fromCol + this.Position));
                }
                else
                {
                    n = "Column" + (this.Position+1).ToString();
                }
            }
            return n;
        }
        set
        {
            string? v = ConvertUtil.ExcelEncodeString(value);
            this.SetXmlNodeString("@name", v);
            if (this._tbl.ShowHeader)
            {
                object? cellValue = this._tbl.WorkSheet.GetValue(this._tbl.Address._fromRow, this._tbl.Address._fromCol + this.Position);
                if (v.Equals(cellValue?.ToString(),StringComparison.CurrentCultureIgnoreCase)==false)
                {
                    this._tbl.WorkSheet.SetValue(this._tbl.Address._fromRow, this._tbl.Address._fromCol + this.Position, value);
                }
            }

            this._tbl.WorkSheet.SetTableTotalFunction(this._tbl, this);
        }
    }
    /// <summary>
    /// A string text in the total row
    /// </summary>
    public string TotalsRowLabel
    {
        get
        {
            return this.GetXmlNodeString("@totalsRowLabel");
        }
        set
        {
            this.SetXmlNodeString("@totalsRowLabel", value);
            this._tbl.WorkSheet.SetValueInner(this._tbl.Address._toRow, this._tbl.Address._fromCol+ this.Position, value);
        }
    }
    /// <summary>
    /// Build-in total row functions.
    /// To set a custom Total row formula use the TotalsRowFormula property
    /// <seealso cref="TotalsRowFormula"/>
    /// </summary>
    public RowFunctions TotalsRowFunction
    {
        get
        {
            if (this.GetXmlNodeString("@totalsRowFunction") == "")
            {
                return RowFunctions.None;
            }
            else
            {
                return (RowFunctions)Enum.Parse(typeof(RowFunctions), this.GetXmlNodeString("@totalsRowFunction"), true);
            }
        }
        set
        {
            if (value == RowFunctions.Custom)
            {
                throw new Exception("Use the TotalsRowFormula-property to set a custom table formula");
            }
            string s = value.ToString();
            s = s.Substring(0, 1).ToLower(CultureInfo.InvariantCulture) + s.Substring(1, s.Length - 1);
            this.SetXmlNodeString("@totalsRowFunction", s);
            this._tbl.WorkSheet.SetTableTotalFunction(this._tbl, this);
        }
    }
    const string TOTALSROWFORMULA_PATH = "d:totalsRowFormula";
    /// <summary>
    /// Sets a custom Totals row Formula.
    /// Be carefull with this property since it is not validated. 
    /// <example>
    /// tbl.Columns[9].TotalsRowFormula = string.Format("SUM([{0}])",tbl.Columns[9].Name);
    /// </example>
    /// </summary>
    public string TotalsRowFormula
    {
        get
        {
            return this.GetXmlNodeString(TOTALSROWFORMULA_PATH);
        }
        set
        {
            if(!string.IsNullOrEmpty(value))
            {
                if (value.StartsWith("="))
                {
                    value = value.Substring(1, value.Length - 1);
                }
            }

            this.SetXmlNodeString("@totalsRowFunction", "custom");
            this.SetXmlNodeString(TOTALSROWFORMULA_PATH, value);
            this._tbl.WorkSheet.SetTableTotalFunction(this._tbl, this);
        }
    }
    const string DATACELLSTYLE_PATH = "@dataCellStyle";
    /// <summary>
    /// The named style for datacells in the column
    /// </summary>
    public string DataCellStyleName
    {
        get
        {
            return this.GetXmlNodeString(DATACELLSTYLE_PATH);
        }
        set
        {
            if(this._tbl.WorkSheet.Workbook.Styles.NamedStyles.FindIndexById(value)<0)
            {
                throw new Exception(string.Format("Named style {0} does not exist.",value));
            }

            this.SetXmlNodeString(this.TopNode, DATACELLSTYLE_PATH, value,true);
               
            int fromRow= this._tbl.Address._fromRow + (this._tbl.ShowHeader?1:0),
                toRow= this._tbl.Address._toRow - (this._tbl.ShowTotal?1:0),
                col= this._tbl.Address._fromCol+ this.Position;

            if (fromRow <= toRow)
            {
                this._tbl.WorkSheet.Cells[fromRow, col, toRow, col].StyleName = value;
            }
        }
    }
    internal const string CALCULATEDCOLUMNFORMULA_PATH = "d:calculatedColumnFormula";

    ExcelTableSlicer _slicer = null;
    /// <summary>
    /// Returns the slicer attached to a column.
    /// If the column has multiple slicers, the first is returned.
    /// </summary>
    public ExcelTableSlicer Slicer 
    {
        get
        {
            if (this._slicer == null)
            {
                ExcelWorkbook? wb = this._tbl.WorkSheet.Workbook;
                if (wb.ExistsNode($"d:extLst/d:ext[@uri='{ExtLstUris.WorkbookSlicerTableUri}']"))
                {
                    foreach (ExcelWorksheet? ws in wb.Worksheets)
                    {
                        foreach (ExcelDrawing? d in ws.Drawings)
                        {
                            if (d is ExcelTableSlicer s && s.TableColumn == this)
                            {
                                this._slicer = s;
                                return this._slicer;
                            }
                        }
                    }
                }
            }
            return this._slicer;
        }
        internal set
        {
            this._slicer = value;
        }
    }
    /// <summary>
    /// Adds a slicer drawing connected to the column
    /// </summary>
    /// <returns>The table slicer drawing object</returns>
    public ExcelTableSlicer AddSlicer()
    {            
        return this._tbl.WorkSheet.Drawings.AddTableSlicer(this);
    }
    /// <summary>
    /// Sets a calculated column Formula.
    /// Be carefull with this property since it is not validated. 
    /// <example>
    /// tbl.Columns[9].CalculatedColumnFormula = string.Format("SUM(MyDataTable[[#This Row],[{0}]])",tbl.Columns[9].Name);  //Reference within the current row
    /// tbl.Columns[9].CalculatedColumnFormula = string.Format("MyDataTable[[#Headers],[{0}]]",tbl.Columns[9].Name);  //Reference to a column header
    /// tbl.Columns[9].CalculatedColumnFormula = string.Format("MyDataTable[[#Totals],[{0}]]",tbl.Columns[9].Name);  //Reference to a column total        
    /// </example>
    /// </summary>
    public string CalculatedColumnFormula
    {
        get
        {
            return this.GetXmlNodeString(CALCULATEDCOLUMNFORMULA_PATH);
        }
        set
        {
            if (string.IsNullOrEmpty(value))
            {
                this.RemoveFormulaNode();
                this.SetTableFormula(true);
            }
            else
            {
                if (value.StartsWith("="))
                {
                    value = value.Substring(1, value.Length - 1);
                }

                this.SetFormula(value);
                this.SetTableFormula(false);
            }
        }
    }
    internal void SetFormula(string formula)
    {
        this.SetXmlNodeString(CALCULATEDCOLUMNFORMULA_PATH, formula);
    }
    internal void RemoveFormulaNode()
    {
        this.DeleteNode(CALCULATEDCOLUMNFORMULA_PATH);
    }

    /// <summary>
    /// The <see cref="ExcelTable"/> containing the table column
    /// </summary>
    public ExcelTable Table
    {
        get
        {
            return this._tbl;
        }
    }
    internal void SetTableFormula(bool clear)
    {
        int fromRow = this._tbl.ShowHeader ? this._tbl.Address._fromRow + 1 : this._tbl.Address._fromRow;
        int toRow = this._tbl.ShowTotal ? this._tbl.Address._toRow - 1 : this._tbl.Address._toRow;
        int colNum = this._tbl.Address._fromCol + this.Position;
        if(clear)
        {
            this._tbl.WorkSheet.Cells[fromRow, colNum, toRow, colNum].Clear();
        }
        else
        {
            this.SetFormulaCells(fromRow, toRow, colNum);
        }
    }

    internal void SetFormulaCells(int fromRow, int toRow, int colNum)
    {
        string r1c1Formula = ExcelCellBase.TranslateToR1C1(this.CalculatedColumnFormula, this._tbl.ShowHeader ? this._tbl.Address._fromRow + 1 : this._tbl.Address._fromRow, colNum);
        bool needsTranslation = r1c1Formula != this.CalculatedColumnFormula;

        ExcelWorksheet? ws = this._tbl.WorkSheet;
        for (int row = fromRow; row <= toRow; row++)
        {
            if(needsTranslation)
            {
                string? f = ExcelCellBase.TranslateFromR1C1(r1c1Formula, row, colNum);
                ws.SetFormula(row, colNum, f);
            }
            else if(ws._formulas.Exists(row, colNum)==false)
            {
                ws.SetFormula(row, colNum, this.CalculatedColumnFormula);
            }
        }
    }
}