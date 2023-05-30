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

using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Table;
using System;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.ExternalReferences;

namespace OfficeOpenXml.FormulaParsing;

/// <summary>
/// EPPlus implementation of the ExcelDataProvider abstract class.
/// </summary>
internal class EpplusExcelDataProvider : ExcelDataProvider
{
    /// <summary>
    /// EPPlus implementation of the <see cref="IRangeInfo"/> interface
    /// </summary>
    public class RangeInfo : IRangeInfo
    {
        internal ExcelWorksheet _ws;
        CellStoreEnumerator<ExcelValue> _values;

        int _fromRow,
            _toRow,
            _fromCol,
            _toCol;

        int _cellCount;
        ExcelAddressBase _address;
        ICellInfo _cell;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="ws">The worksheet</param>
        /// <param name="fromRow"></param>
        /// <param name="fromCol"></param>
        /// <param name="toRow"></param>
        /// <param name="toCol"></param>
        public RangeInfo(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ExcelAddressBase? address = new ExcelAddressBase(fromRow, fromCol, toRow, toCol);
            address._ws = ws.Name;
            this.SetAddress(ws, address);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="address"></param>
        public RangeInfo(ExcelWorksheet ws, ExcelAddressBase address)
        {
            this.SetAddress(ws, address);
        }

        private void SetAddress(ExcelWorksheet ws, ExcelAddressBase address)
        {
            this._ws = ws;
            this._fromRow = address._fromRow;
            this._fromCol = address._fromCol;
            this._toRow = address._toRow;
            this._toCol = address._toCol;
            this._address = address;

            if (this._ws != null && this._ws.IsDisposed == false)
            {
                this._values = new CellStoreEnumerator<ExcelValue>(this._ws._values, this._fromRow, this._fromCol, this._toRow, this._toCol);
                this._cell = new CellInfo(this._ws, this._values);
            }
        }

        /// <summary>
        /// The total number of cells (including empty) of the range
        /// </summary>
        /// <returns></returns>
        public int GetNCells()
        {
            return (this._toRow - this._fromRow + 1) * (this._toCol - this._fromCol + 1);
        }

        /// <summary>
        /// Returns true if the range represents a reference
        /// </summary>
        public bool IsRef
        {
            get { return this._ws == null || this._fromRow < 0 || this._toRow < 0; }
        }

        /// <summary>
        /// Returns true if the range is empty
        /// </summary>
        public bool IsEmpty
        {
            get
            {
                if (this._cellCount > 0)
                {
                    return false;
                }
                else if (this._values == null)
                {
                    return true;
                }
                else if (this._values.Next())
                {
                    this._values.Reset();

                    return false;
                }
                else
                {
                    return true;
                }
            }
        }

        /// <summary>
        /// Returns true if more than one cell
        /// </summary>
        public bool IsMulti
        {
            get
            {
                if (this._cellCount == 0)
                {
                    if (this._values == null)
                    {
                        return false;
                    }

                    if (this._values.Next() && this._values.Next())
                    {
                        this._values.Reset();

                        return true;
                    }
                    else
                    {
                        this._values.Reset();

                        return false;
                    }
                }
                else if (this._cellCount > 1)
                {
                    return true;
                }

                return false;
            }
        }

        /// <summary>
        /// Current cell
        /// </summary>
        public ICellInfo Current
        {
            get { return this._cell; }
        }

        /// <summary>
        /// The worksheet
        /// </summary>
        public ExcelWorksheet Worksheet
        {
            get { return this._ws; }
        }

        /// <summary>
        /// Runs at dispose of this instance
        /// </summary>
        public void Dispose()
        {
            //_values = null;
            //_ws = null;
            //_cell = null;
        }

        /// <summary>
        /// IEnumerator.Current
        /// </summary>
        object System.Collections.IEnumerator.Current
        {
            get { return this; }
        }

        /// <summary>
        /// Moves to next cell
        /// </summary>
        /// <returns></returns>
        public bool MoveNext()
        {
            if (this._values == null)
            {
                return false;
            }

            this._cellCount++;

            return this._values.MoveNext();
        }

        /// <summary>
        /// 
        /// </summary>
        public void Reset()
        {
            this._cellCount = 0;
            this._values?.Init();
        }

        /// <summary>
        /// Moves to next cell
        /// </summary>
        /// <returns></returns>
        public bool NextCell()
        {
            if (this._values == null)
            {
                return false;
            }

            this._cellCount++;

            return this._values.MoveNext();
        }

        /// <summary>
        /// Returns enumerator for cells
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ICellInfo> GetEnumerator()
        {
            this.Reset();

            return this;
        }

        /// <summary>
        /// Returns enumerator for cells
        /// </summary>
        /// <returns></returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this;
        }

        /// <summary>
        /// Address of the range
        /// </summary>
        public ExcelAddressBase Address
        {
            get { return this._address; }
        }

        /// <summary>
        /// Returns the cell value 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public object GetValue(int row, int col)
        {
            return this._ws?.GetValue(row, col);
        }

        public object GetOffset(int rowOffset, int colOffset)
        {
            if (this._values == null)
            {
                return null;
            }

            if (this._values.Row < this._fromRow || this._values.Column < this._fromCol)
            {
                return this._ws.GetValue(this._fromRow + rowOffset, this._fromCol + colOffset);
            }
            else
            {
                return this._ws.GetValue(this._values.Row + rowOffset, this._values.Column + colOffset);
            }
        }
    }

    public class CellInfo : ICellInfo
    {
        ExcelWorksheet _ws;
        CellStoreEnumerator<ExcelValue> _values;

        internal CellInfo(ExcelWorksheet ws, CellStoreEnumerator<ExcelValue> values)
        {
            this._ws = ws;
            this._values = values;
        }

        public string Address
        {
            get { return this._values.CellAddress; }
        }

        public int Row
        {
            get { return this._values.Row; }
        }

        public int Column
        {
            get { return this._values.Column; }
        }

        public string Formula
        {
            get { return this._ws.GetFormula(this._values.Row, this._values.Column); }
        }

        public object Value
        {
            get
            {
                if (this._ws._flags.GetFlagValue(this._values.Row, this._values.Column, CellFlags.RichText))
                {
                    return this._ws.GetRichText(this._values.Row, this._values.Column, null).Text;
                }
                else
                {
                    return this._values.Value._value;
                }
            }
        }

        public double ValueDouble
        {
            get { return ConvertUtil.GetValueDouble(this._values.Value._value, true); }
        }

        public double ValueDoubleLogical
        {
            get { return ConvertUtil.GetValueDouble(this._values.Value._value, false); }
        }

        public bool IsHiddenRow
        {
            get
            {
                RowInternal? row = this._ws.GetValueInner(this._values.Row, 0) as RowInternal;

                if (row != null)
                {
                    return row.Hidden || row.Height == 0;
                }
                else
                {
                    return false;
                }
            }
        }

        public bool IsExcelError
        {
            get { return ExcelErrorValue.Values.IsErrorValue(this._values.Value._value); }
        }

        public IList<Token> Tokens
        {
            get { return this._ws._formulaTokens.GetValue(this._values.Row, this._values.Column); }
        }

        public ulong Id
        {
            get { return ExcelCellBase.GetCellId(this._ws.IndexInList, this._values.Row, this._values.Column); }
        }

        public string WorksheetName
        {
            get { return this._ws.Name; }
        }
    }

    public class NameInfo : INameInfo
    {
        public ulong Id { get; set; }

        public string Worksheet { get; set; }

        public string Name { get; set; }

        public string Formula { get; set; }

        public IList<Token> Tokens { get; internal set; }

        public object Value { get; set; }
    }

    private readonly ExcelPackage _package;
    private ExcelWorksheet _currentWorksheet;
    //private RangeAddressFactory _rangeAddressFactory;
    private Dictionary<ulong, INameInfo> _names = new Dictionary<ulong, INameInfo>();

    internal EpplusExcelDataProvider()
        : this(new ExcelPackage())
    {
    }

    public EpplusExcelDataProvider(ExcelPackage package)
    {
        this._package = package;

        //this._rangeAddressFactory = new RangeAddressFactory(this);
    }

    public override IEnumerable<string> GetWorksheets()
    {
        return this._package.Workbook.Worksheets.Select(x => x.Name);
    }

    public override ExcelNamedRangeCollection GetWorksheetNames(string worksheet)
    {
        ExcelWorksheet? ws = this._package.Workbook.Worksheets[worksheet];

        if (ws != null)
        {
            return ws.Names;
        }
        else
        {
            return null;
        }
    }

    public override int GetWorksheetIndex(string worksheetName)
    {
        for (int ix = 1; ix <= this._package.Workbook.Worksheets.Count; ix++)
        {
            ExcelWorksheet? ws = this._package.Workbook.Worksheets[ix - 1];

            if (string.Compare(worksheetName, ws.Name, true) == 0)
            {
                return ix;
            }
        }

        return -1;
    }

    public override ExcelTable GetExcelTable(string name)
    {
        foreach (ExcelWorksheet? ws in this._package.Workbook.Worksheets)
        {
            if (ws is ExcelChartsheet)
            {
                continue;
            }

            if (ws.Tables._tableNames.ContainsKey(name))
            {
                return ws.Tables[name];
            }
        }

        return null;
    }

    public override ExcelNamedRangeCollection GetWorkbookNameValues()
    {
        return this._package.Workbook.Names;
    }

    public override IRangeInfo GetRange(string worksheet, int fromRow, int fromCol, int toRow, int toCol)
    {
        this.SetCurrentWorksheet(worksheet);
        string? wsName = string.IsNullOrEmpty(worksheet) ? this._currentWorksheet.Name : worksheet;
        ExcelWorksheet? ws = this._package.Workbook.Worksheets[wsName];

        if (ws == null)
        {
            throw new ExcelErrorValueException(eErrorType.Ref);
        }
        else
        {
            return new RangeInfo(ws, fromRow, fromCol, toRow, toCol);
        }
    }

    public override IRangeInfo GetRange(string worksheet, int row, int column, string address)
    {
        this.SetCurrentWorksheet(worksheet);
        ExcelAddressBase? addr = new ExcelAddressBase(address, this._package.Workbook, worksheet);

        if (addr.Table != null && string.IsNullOrEmpty(addr._wb))
        {
            addr = ConvertToA1C1(this._package, addr, new ExcelAddressBase(row, column, row, column));
        }

        return this.GetRangeInternal(addr);
    }

    public override IRangeInfo GetRange(string worksheet, string address)
    {
        this.SetCurrentWorksheet(worksheet);
        ExcelAddressBase? addr = new ExcelAddressBase(address, this._package.Workbook, worksheet);

        if (addr.Table != null)
        {
            addr = ConvertToA1C1(this._package, addr, addr);
        }

        return this.GetRangeInternal(addr);
    }

    private IRangeInfo GetRangeInternal(ExcelAddressBase addr)
    {
        if (addr.IsExternal)
        {
            return GetExternalRangeInfo(addr, addr.WorkSheetName, this._package.Workbook);
        }
        else
        {
            string? wsName = string.IsNullOrEmpty(addr.WorkSheetName) ? this._currentWorksheet.Name : addr.WorkSheetName;
            ExcelWorksheet? ws = this._package.Workbook.Worksheets[wsName];

            if (ws == null)
            {
                throw new ExcelErrorValueException(eErrorType.Ref);
            }

            return new RangeInfo(ws, addr);
        }
    }

    private static IRangeInfo GetExternalRangeInfo(ExcelAddressBase addr, string wsName, ExcelWorkbook wb)
    {
        ExcelExternalWorkbook externalWb;
        int ix = wb.ExternalLinks.GetExternalLink(addr._wb);

        if (ix >= 0)
        {
            externalWb = wb.ExternalLinks[ix].As.ExternalWorkbook;
        }
        else
        {
            throw new ExcelErrorValueException(eErrorType.Ref);
        }

        if (externalWb?.Package == null)
        {
            if (addr.Table != null)
            {
                throw new ExcelErrorValueException(eErrorType.Ref);
            }

            return new EpplusExcelExternalRangeInfo(externalWb, wb, addr);
        }
        else
        {
            addr = addr.ToInternalAddress();
            ExcelWorksheet ws;

            if (addr.Table == null)
            {
                ws = externalWb.Package.Workbook.Worksheets[wsName];
            }
            else
            {
                addr = ConvertToA1C1(externalWb.Package, addr, addr);
                ws = externalWb.Package.Workbook.Worksheets[addr.WorkSheetName];
            }

            return new RangeInfo(ws, addr);
        }
    }

    private static ExcelAddress ConvertToA1C1(ExcelPackage package, ExcelAddressBase addr, ExcelAddressBase refAddress)
    {
        //Convert the Table-style Address to an A1C1 address
        addr.SetRCFromTable(package, refAddress);
        ExcelAddress? a = new ExcelAddress(addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
        a._ws = addr._ws;

        return a;
    }

    public override INameInfo GetName(string worksheet, string name)
    {
        if (ExcelCellBase.IsExternalAddress(name))
        {
            return this.GetExternalName(name);
        }
        else
        {
            return this.GetLocalName(this._package, worksheet, name);
        }
    }

    private INameInfo GetExternalName(string name)
    {
        string? extRef = ExcelCellBase.GetWorkbookFromAddress(name);
        int ix = this._package.Workbook.ExternalLinks.GetExternalLink(extRef);

        if (ix >= 0)
        {
            ExcelExternalWorkbook? externalWorkbook = this._package.Workbook.ExternalLinks[ix].As.ExternalWorkbook;

            if (externalWorkbook != null)
            {
                if (externalWorkbook.Package == null)
                {
                    return GetNameFromCache(externalWorkbook, name);
                }
                else
                {
                    name = name.Substring(name.IndexOf("]") + 1);

                    if (name.StartsWith("!"))
                    {
                        name = name.Substring(1);
                    }

                    return this.GetLocalName(externalWorkbook.Package, "", name);
                }
            }
        }

        return null;
    }

    private static INameInfo GetNameFromCache(ExcelExternalWorkbook externalWorkbook, string name)
    {
        ExcelExternalDefinedName nameItem;

        int ix = -1;
        string? sheetName = ExcelAddressBase.GetWorksheetPart(name, "", ref ix);

        if (string.IsNullOrEmpty(sheetName))
        {
            if (ix > 0)
            {
                name = name.Substring(ix);
            }

            nameItem = externalWorkbook.CachedNames[name];
        }
        else
        {
            if (ix >= 0)
            {
                name = name.Substring(ix);
            }

            nameItem = externalWorkbook.CachedWorksheets[sheetName].CachedNames[name];
        }

        object value;

        if (!string.IsNullOrEmpty(nameItem.RefersTo))
        {
            string? nameAddress = nameItem.RefersTo.TrimStart('=');
            ExcelAddressBase address = new ExcelAddressBase(nameAddress);

            if (address.Address == "#REF!")
            {
                value = ExcelErrorValue.Create(eErrorType.Ref);
            }
            else
            {
                value = new EpplusExcelExternalRangeInfo(externalWorkbook, null, address);
            }
        }
        else
        {
            value = ExcelErrorValue.Create(eErrorType.Name);
        }

        return new NameInfo() { Name = name, Value = value };
    }

    private INameInfo GetLocalName(ExcelPackage package, string worksheet, string name)
    {
        ExcelNamedRange nameItem;
        ExcelWorksheet ws;
        int ix = name.IndexOf('!');

        if (ix > 0)
        {
            string? wsName = ExcelAddressBase.GetWorksheetPart(name, worksheet, ref ix);

            if (!string.IsNullOrEmpty(wsName))
            {
                name = name.Substring(ix);
                worksheet = wsName;
            }
        }

        if (string.IsNullOrEmpty(worksheet))
        {
            if (package._workbook.Names.ContainsKey(name))
            {
                nameItem = package._workbook.Names[name];
            }
            else
            {
                return null;
            }

            ws = null;
        }
        else
        {
            ws = package._workbook.Worksheets[worksheet];

            if (ws != null && ws.Names.ContainsKey(name))
            {
                nameItem = ws.Names[name];
            }
            else if (package._workbook.Names.ContainsKey(name))
            {
                nameItem = package._workbook.Names[name];
            }
            else
            {
                ExcelTable? tbl = ws.Tables[name];

                if (tbl != null)
                {
                    nameItem = new ExcelNamedRange(name, ws, ws, tbl.DataRange.Address, -1);
                }
                else
                {
                    ExcelWorksheet? wsName = package.Workbook.Worksheets[name];

                    if (wsName == null)
                    {
                        return null;
                    }

                    nameItem = new ExcelNamedRange(name, ws, wsName, "A:XFD", -1);
                }
            }
        }

        ulong id = ExcelCellBase.GetCellId(nameItem.LocalSheetId, nameItem.Index, 0);

        if (this._names.ContainsKey(id))
        {
            return this._names[id];
        }
        else
        {
            NameInfo? ni = new NameInfo()
            {
                Id = id,
                Name = name,
                Worksheet = string.IsNullOrEmpty(worksheet) ? nameItem.Worksheet == null ? nameItem._ws : nameItem.Worksheet.Name : worksheet,
                Formula = nameItem.Formula
            };

            if (nameItem._fromRow > 0)
            {
                ni.Value = new RangeInfo(nameItem.Worksheet ?? ws, nameItem._fromRow, nameItem._fromCol, nameItem._toRow, nameItem._toCol);
            }
            else
            {
                ni.Value = nameItem.Value;
            }

            this._names.Add(id, ni);

            return ni;
        }
    }

    public override IEnumerable<object> GetRangeValues(string address)
    {
        this.SetCurrentWorksheet(ExcelAddressInfo.Parse(address));
        ExcelAddress? addr = new ExcelAddress(address);
        string? wsName = string.IsNullOrEmpty(addr.WorkSheetName) ? this._currentWorksheet.Name : addr.WorkSheetName;
        ExcelWorksheet? ws = this._package.Workbook.Worksheets[wsName];

        return (IEnumerable<object>)new CellStoreEnumerator<ExcelValue>(ws._values, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
    }

    public object GetValue(int row, int column)
    {
        return this._currentWorksheet.GetValueInner(row, column);
    }

    public bool IsMerged(int row, int column)
    {
        //return _currentWorksheet._flags.GetFlagValue(row, column, CellFlags.Merged);
        return this._currentWorksheet.MergedCells[row, column] != null;
    }

    public bool IsHidden(int row, int column)
    {
        return this._currentWorksheet.Column(column).Hidden
               || this._currentWorksheet.Column(column).Width == 0
               || this._currentWorksheet.Row(row).Hidden
               || this._currentWorksheet.Row(column).Height == 0;
    }

    public override object GetCellValue(string sheetName, int row, int col)
    {
        this.SetCurrentWorksheet(sheetName);

        return this._currentWorksheet.GetValueInner(row, col);
    }

    public override ulong GetCellId(string sheetName, int row, int col)
    {
        if (string.IsNullOrEmpty(sheetName))
        {
            return 0;
        }

        ExcelWorksheet? worksheet = this._package.Workbook.Worksheets[sheetName];
        int wsIx = worksheet != null ? worksheet.IndexInList : 0;

        return ExcelCellBase.GetCellId(wsIx, row, col);
    }

    public override ExcelCellAddress GetDimensionEnd(string worksheet)
    {
        ExcelCellAddress address = null;

        try
        {
            address = this._package.Workbook.Worksheets[worksheet].Dimension.End;
        }
        catch
        {
        }

        return address;
    }

    private void SetCurrentWorksheet(ExcelAddressInfo addressInfo)
    {
        if (addressInfo.WorksheetIsSpecified)
        {
            this._currentWorksheet = this._package.Workbook.Worksheets[addressInfo.Worksheet];
        }
        else if (this._currentWorksheet == null)
        {
            this._currentWorksheet = this._package.Workbook.Worksheets.First();
        }
    }

    private void SetCurrentWorksheet(string worksheetName)
    {
        if (!string.IsNullOrEmpty(worksheetName))
        {
            this._currentWorksheet = this._package.Workbook.Worksheets[worksheetName];
        }
        else
        {
            this._currentWorksheet = this._package.Workbook.Worksheets.First();
        }
    }

    //public override void SetCellValue(string address, object value)
    //{
    //    var addressInfo = ExcelAddressInfo.Parse(address);
    //    var ra = _rangeAddressFactory.Create(address);
    //    SetCurrentWorksheet(addressInfo);
    //    //var valueInfo = (ICalcEngineValueInfo)_currentWorksheet;
    //    //valueInfo.SetFormulaValue(ra.FromRow + 1, ra.FromCol + 1, value);
    //    _currentWorksheet.Cells[ra.FromRow + 1, ra.FromCol + 1].Value = value;
    //}

    public override void Dispose()
    {
        this._package.Dispose();
    }

    public override int ExcelMaxColumns
    {
        get { return ExcelPackage.MaxColumns; }
    }

    public override int ExcelMaxRows
    {
        get { return ExcelPackage.MaxRows; }
    }

    public override string GetRangeFormula(string worksheetName, int row, int column)
    {
        this.SetCurrentWorksheet(worksheetName);

        return this._currentWorksheet.GetFormula(row, column);
    }

    public override object GetRangeValue(string worksheetName, int row, int column)
    {
        this.SetCurrentWorksheet(worksheetName);

        return this._currentWorksheet.GetValue(row, column);
    }

    public override string GetFormat(object value, string format)
    {
        ExcelStyles? styles = this._package.Workbook.Styles;
        ExcelNumberFormatXml.ExcelFormatTranslator ft = null;

        foreach (ExcelNumberFormatXml? f in styles.NumberFormats)
        {
            if (f.Format == format)
            {
                ft = f.FormatTranslator;

                break;
            }
        }

        ft ??= new ExcelNumberFormatXml.ExcelFormatTranslator(format, -1);

        return ValueToTextHandler.FormatValue(value, false, ft, null);
    }

    public override List<Token> GetRangeFormulaTokens(string worksheetName, int row, int column)
    {
        return this._package.Workbook.Worksheets[worksheetName]._formulaTokens.GetValue(row, column);
    }

    public override bool IsRowHidden(string worksheetName, int row)
    {
        bool b = this._package.Workbook.Worksheets[worksheetName].Row(row).Height == 0 || this._package.Workbook.Worksheets[worksheetName].Row(row).Hidden;

        return b;
    }

    public override void Reset()
    {
        this._names = new Dictionary<ulong, INameInfo>(); //Reset name cache.            
    }

    public override bool IsExternalName(string name)
    {
        if (name[0] != '[')
        {
            return false;
        }

        int ixEnd = name.IndexOf("]");

        if (ixEnd > 0)
        {
            string? ix = name.Substring(1, ixEnd - 1);
            int extRef = this._package.Workbook.ExternalLinks.GetExternalLink(ix);

            if (extRef < 0)
            {
                return false;
            }

            ExcelExternalWorkbook? extBook = this._package.Workbook.ExternalLinks[extRef].As.ExternalWorkbook;

            if (extBook == null)
            {
                return false;
            }

            string? address = name.Substring(ixEnd + 1);

            if (address.StartsWith("!"))
            {
                return extBook.CachedNames.ContainsKey(address.Substring(1));
            }
            else
            {
                int addressStart = -1;
                string? sheetName = ExcelAddressBase.GetWorksheetPart(address, "", ref addressStart);

                if (extBook.CachedWorksheets.ContainsKey(sheetName) && addressStart > 0)
                {
                    return extBook.CachedWorksheets[sheetName].CachedNames.ContainsKey(address.Substring(addressStart));
                }
            }
        }

        return false;
    }

    //public override void SetToTableAddress(ExcelAddress address)
    //{
    //    address.SetRCFromTable(_package, address);
    //}
}