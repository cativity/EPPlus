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
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.Table;

namespace OfficeOpenXml;

/// <summary>
/// A range address
/// </summary>
/// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
public class ExcelAddressBase : ExcelCellBase
{
    internal int _fromRow=-1, _toRow, _fromCol, _toCol;
    internal bool _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed;
    internal string _wb;
    internal string _ws;
    internal string _address;

    internal enum eAddressCollition
    {
        No,
        Partly,
        Inside,
        Equal
    }
    #region "Constructors"
    internal ExcelAddressBase()
    {
    }
    /// <summary>
    /// Creates an Address object
    /// </summary>
    /// <param name="fromRow">start row</param>
    /// <param name="fromCol">start column</param>
    /// <param name="toRow">End row</param>
    /// <param name="toColumn">End column</param>
    public ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn)
    {
        this._fromRow = fromRow;
        this._toRow = toRow;
        this._fromCol = fromCol;
        this._toCol = toColumn;
        this.Validate();

        this._address = GetAddress(this._fromRow, this._fromCol, this._toRow, this._toCol);
    }
    /// <summary>
    /// Creates an Address object
    /// </summary>
    /// <param name="worksheetName">Worksheet name</param>
    /// <param name="fromRow">Start row</param>
    /// <param name="fromCol">Start column</param>
    /// <param name="toRow">End row</param>
    /// <param name="toColumn">End column</param>
    public ExcelAddressBase(string worksheetName, int fromRow, int fromCol, int toRow, int toColumn)
    {
        this._ws = worksheetName;
        this._fromRow = fromRow;
        this._toRow = toRow;
        this._fromCol = fromCol;
        this._toCol = toColumn;
        this.Validate();

        this._address = GetAddress(this._fromRow, this._fromCol, this._toRow, this._toCol);
    }

    internal static bool IsTableAddress(string address)
    {
        SplitAddress(address, out string wb, out string ws, out string intAddress);
        int lPos = intAddress.IndexOf('[');
        if(lPos >= 0) 
        {
            int rPos= intAddress.IndexOf(']',lPos);
            if(rPos>lPos)
            {
                char c=intAddress[lPos+1];
                return !((c >= '0' && c <= '9') || c == '-');
            }
        }
        return false;
    }

    /// <summary>
    /// Creates an Address object
    /// </summary>
    /// <param name="fromRow">Start row</param>
    /// <param name="fromCol">Start column</param>
    /// <param name="toRow">End row</param>
    /// <param name="toColumn">End column</param>
    /// <param name="fromRowFixed">Start row fixed</param>
    /// <param name="fromColFixed">Start column fixed</param>
    /// <param name="toRowFixed">End row fixed</param>
    /// <param name="toColFixed">End column fixed</param>
    public ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn, bool fromRowFixed, bool fromColFixed, bool toRowFixed, bool toColFixed) :
        this(fromRow, fromCol, toRow, toColumn, fromRowFixed, fromColFixed, toRowFixed, toColFixed, null, null)
    {

    }
    internal ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn, bool fromRowFixed, bool fromColFixed, bool toRowFixed, bool toColFixed, string worksheetName, string prevAddress)
    {
        this._fromRow = fromRow;
        this._toRow = toRow;
        this._fromCol = fromCol;
        this._toCol = toColumn;
        this._fromRowFixed = fromRowFixed;
        this._fromColFixed = fromColFixed;
        this._toRowFixed = toRowFixed;
        this._toColFixed = toColFixed;
        this._ws = worksheetName;
        this.Validate();
        bool prevAddressHasWs = prevAddress != null && prevAddress.IndexOf("!") > 0 && !prevAddress.EndsWith("!");
        this._address = GetAddress(this._fromRow, this._fromCol, this._toRow, this._toCol, this._fromRowFixed, fromColFixed, this._toRowFixed, this._toColFixed );
        if(prevAddressHasWs && !string.IsNullOrEmpty(this._ws))
        {
            if(ExcelWorksheet.NameNeedsApostrophes(this._ws))
            {
                this._address = $"'{this._ws.Replace("'","''")}'!{this._address}";
            }
            else
            {
                this._address = $"{this._ws}!{this._address}";
            }
        }
    }
    /// <summary>
    /// Creates an Address object
    /// </summary>
    /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
    /// <param name="address">The Excel Address</param>
    /// <param name="wb">The workbook to verify any defined names from</param>
    /// <param name="wsName">The name of the worksheet the address referes to</param>
    /// <ws></ws>
    public ExcelAddressBase(string address, ExcelWorkbook wb=null, string wsName=null)
    {
        this.SetAddress(address, wb, wsName);
        if (string.IsNullOrEmpty(this._ws) && string.IsNullOrEmpty(this._wb))
        {
            this._ws = wsName;
        }
    }
    /// <summary>
    /// Creates an Address object
    /// </summary>
    /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
    /// <param name="address">The Excel Address</param>
    /// <param name="pck">Reference to the package to find information about tables and names</param>
    /// <param name="referenceAddress">The address</param>
    public ExcelAddressBase(string address, ExcelPackage pck, ExcelAddressBase referenceAddress)
    {
        this.SetAddress(address, null, null);
        this.SetRCFromTable(pck, referenceAddress);
    }

    internal void SetRCFromTable(ExcelPackage pck, ExcelAddressBase referenceAddress)
    {
        if (string.IsNullOrEmpty(this._wb) && this.Table != null)
        {
            foreach (ExcelWorksheet? ws in pck.Workbook.Worksheets)
            {
                if (ws is ExcelChartsheet)
                {
                    continue;
                }

                foreach (ExcelTable? t in ws.Tables)
                {
                    if (t.Name.Equals(this.Table.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        this._ws = ws.Name;
                        if (this.Table.IsAll)
                        {
                            this._fromRow = t.Address._fromRow;
                            this._toRow = t.Address._toRow;
                        }
                        else
                        {
                            if (this.Table.IsThisRow)
                            {
                                if (referenceAddress == null)
                                {
                                    this._fromRow = -1;
                                    this._toRow = -1;
                                }
                                else
                                {
                                    this._fromRow = referenceAddress._fromRow;
                                    this._toRow = this._fromRow;
                                }
                            }
                            else if (this.Table.IsHeader && this.Table.IsData)
                            {
                                this._fromRow = t.Address._fromRow;
                                this._toRow = t.ShowTotal ? t.Address._toRow - 1 : t.Address._toRow;
                            }
                            else if (this.Table.IsData && this.Table.IsTotals)
                            {
                                this._fromRow = t.ShowHeader ? t.Address._fromRow + 1 : t.Address._fromRow;
                                this._toRow = t.Address._toRow;
                            }
                            else if (this.Table.IsHeader)
                            {
                                this._fromRow = t.ShowHeader ? t.Address._fromRow : -1;
                                this._toRow = t.ShowHeader ? t.Address._fromRow : -1;
                            }
                            else if (this.Table.IsTotals)
                            {
                                this._fromRow = t.ShowTotal ? t.Address._toRow : -1;
                                this._toRow = t.ShowTotal ? t.Address._toRow : -1;
                            }
                            else
                            {
                                this._fromRow = t.ShowHeader ? t.Address._fromRow + 1 : t.Address._fromRow;
                                this._toRow = t.ShowTotal ? t.Address._toRow - 1 : t.Address._toRow;
                            }
                        }

                        if (string.IsNullOrEmpty(this.Table.ColumnSpan))
                        {
                            this._fromCol = t.Address._fromCol;
                            this._toCol = t.Address._toCol;
                            return;
                        }
                        else
                        {
                            int col = t.Address._fromCol;
                            string[]? cols = this.Table.ColumnSpan.Split(':');
                            foreach (ExcelTableColumn? c in t.Columns)
                            {
                                if (this._fromCol <= 0 && cols[0].Equals(c.Name, StringComparison.OrdinalIgnoreCase))   //Issue15063 Add invariant igore case
                                {
                                    this._fromCol = col;
                                    if (cols.Length == 1)
                                    {
                                        this._toCol = this._fromCol;
                                        return;
                                    }
                                }
                                else if (cols.Length > 1 && this._fromCol > 0 && cols[1].Equals(c.Name, StringComparison.OrdinalIgnoreCase)) //Issue15063 Add invariant igore case
                                {
                                    this._toCol = col;
                                    return;
                                }

                                col++;
                            }
                        }
                    }
                }
            }
        }
    }
    internal string ChangeTableName(string prevName, string name)
    {
        if (this.LocalAddress.StartsWith(prevName +"[", StringComparison.CurrentCultureIgnoreCase))
        {
            string? wsPart = "";
            int ix = this._address.TrimEnd().LastIndexOf('!', this._address.Length - 2);  //Last index can be ! if address is #REF!, so check from                 
            if (ix >= 0)
            {
                wsPart= this._address.Substring(0, ix);
            }

            return wsPart + name + this.LocalAddress.Substring(prevName.Length);
        }
        else
        {
            return this._address;
        }
    }
    internal ExcelAddressBase Intersect(ExcelAddressBase address)
    {
        if(address._fromRow > this._toRow || this._toRow < address._fromRow ||
           address._fromCol > this._toCol || this._toCol < address._fromCol || this._fromRow > address._toRow || address._toRow < this._fromRow || this._fromCol > address._toCol || address._toCol < this._fromCol
          )
        {
            return null;
        }
            
        int fromRow = Math.Max(address._fromRow, this._fromRow);
        int toRow = Math.Min(address._toRow, this._toRow);
        int fromCol = Math.Max(address._fromCol, this._fromCol);
        int toCol = Math.Min(address._toCol, this._toCol);

        return new ExcelAddressBase(fromRow, fromCol, toRow, toCol);
    }
    /// <summary>
    /// Returns the parts of this address that not intersects with <paramref name="address"/>
    /// </summary>
    /// <param name="address">The address to intersect with</param>
    /// <returns>The addresses not intersecting with <paramref name="address"/></returns>
    internal ExcelAddressBase IntersectReversed(ExcelAddressBase address)
    {
        if (address._fromRow > this._toRow || this._toRow < address._fromRow ||
            address._fromCol > this._toCol || this._toCol < address._fromCol || this._fromRow > address._toRow || address._toRow < this._fromRow || this._fromCol > address._toCol || address._toCol < this._fromCol ||
            (string.IsNullOrEmpty(address._ws) == false && string.IsNullOrEmpty(this._ws) == false && address._ws != this._ws))
        {
            return this;
        }
        string retAddress = "";
        int fromRow = this._fromRow, fromCol = this._fromCol, toCol = this._toCol;

        if (this._fromCol < address._fromCol)
        {
            retAddress = GetAddress(fromRow, fromCol, this._toRow, address._fromCol - 1) + ",";
            fromCol = address._fromCol;
        }

        if (this._fromRow < address._fromRow)
        {
            retAddress += GetAddress(fromRow, fromCol, address._fromRow - 1, toCol) + ",";
            fromRow = address._fromRow;
        }

        if (this._toCol > address._toCol)
        {
            retAddress += GetAddress(fromRow, address._toCol + 1, this._toRow, toCol) + ",";
            toCol = address._toCol;
        }

        if (this._toRow > address._toRow)
        {
            retAddress += GetAddress(address._toRow + 1, fromCol, this._toRow, toCol) + ",";
        }
        return string.IsNullOrEmpty(retAddress) ? null : new ExcelAddressBase(retAddress.Substring(0, retAddress.Length - 1));
    }

    internal bool IsInside(ExcelAddressBase effectedAddress)
    {
        eAddressCollition c = this.Collide(effectedAddress);
        return c == eAddressCollition.Equal ||
               c == eAddressCollition.Inside;
    }
    /// <summary>
    /// Address is an defined name
    /// </summary>
    /// <param name="address">the name</param>
    /// <param name="isName">Should always be true</param>
    internal ExcelAddressBase(string address, bool isName)
    {
        if (isName)
        {
            this._address = address;
            this._fromRow = -1;
            this._fromCol = -1;
            this._toRow = -1;
            this._toCol = -1;
            this._start = null;
            this._end = null;
        }
        else
        {
            this.SetAddress(address, null, null);
        }
    }
    /// <summary>
    /// Sets the address
    /// </summary>
    /// <param name="address">The address</param>
    /// <param name="wb"></param>
    /// <param name="wsName"></param>
    protected internal void SetAddress(string address, ExcelWorkbook wb, string wsName)
    {
        address = address.Trim();
        if (address.Length > 0 && (address[0] == '\'' || address[0] == '['))
        {
            this.SetWbWs(address);
        }
        else
        {
            this._address = address;
        }

        this._addresses = null;
        if (this._address.IndexOfAny(new char[] {',','!', '['}) > -1)
        {
            this._firstAddress = null;
            //Advanced address. Including Sheet or multi or table.
            this.ExtractAddress(this._address);
        }
        else
        {
            //Simple address
            GetRowColFromAddress(this._address, out this._fromRow, out this._fromCol, out this._toRow, out this._toCol, out this._fromRowFixed, out this._fromColFixed,  out this._toRowFixed, out this._toColFixed, wb, wsName);
            this._start = null;
            this._end = null;
        }

        this._address = address;
        this.Validate();
    }

    internal ExcelAddressBase ToInternalAddress()
    {
        if(this._address.StartsWith("["))
        {
            int ix = this._address.IndexOf("]", 1);
            if (ix > 0)
            {
                if(this._address[ix+1]=='!')
                {
                    ix++;
                }
                string? a = this._address.Substring(ix+1);
                    
                return new ExcelAddressBase(a);
            }
            return this;
        }
        else
        {
            return this;
        }
    }

    internal protected virtual void BeforeChangeAddress()
    {
    }

    /// <summary>
    /// Called when the address changes
    /// </summary>
    internal protected virtual void ChangeAddress()
    {
    }
    private void SetWbWs(string address)
    {
        int pos;
        if (address[0] == '[')
        {
            pos = address.IndexOf(']');
            this._wb = address.Substring(1, pos - 1);
            this._ws = address.Substring(pos + 1);                
        }
        else
        {
            this._wb = "";
            this._ws = address;
        }
        if(this._ws.StartsWith("'", StringComparison.OrdinalIgnoreCase))
        {
            pos = this._ws.IndexOf("'",1, StringComparison.OrdinalIgnoreCase);
            while(pos>0 && pos+1< this._ws.Length && this._ws[pos+1]=='\'')
            {
                this._ws = this._ws.Substring(0, pos) + this._ws.Substring(pos+1);
                pos = this._ws.IndexOf("'", pos+1, StringComparison.OrdinalIgnoreCase);
            }
            if (pos>0)
            {
                if(this._ws.Length-1==pos)
                {
                    this._address = "A:XFD";
                }
                else if (this._ws[pos+1]!='!')
                {
                    throw new InvalidOperationException($"Address is not valid {address}. Missing ! after sheet name.");
                }
                else
                {
                    this._address = this._ws.Substring(pos + 2);
                }

                this._ws = this._ws.Substring(1, pos-1);
                if(this._ws.StartsWith("["))
                {
                    int ix = this._ws.IndexOf("]", 1);
                    if(ix>0)
                    {
                        this._wb = this._ws.Substring(1, ix - 1);
                        this._ws = this._ws.Substring(ix+1);
                    }
                }
                pos = this._address.IndexOf(":'", StringComparison.OrdinalIgnoreCase);
                if(pos>0)
                {
                    string? a1 = this._address.Substring(0,pos);
                    pos = this._address.LastIndexOf("\'!", StringComparison.OrdinalIgnoreCase);
                    if (pos > 0)
                    {
                        string? a2 = this._address.Substring(pos+2);
                        this._address=a1 + ":" + a2; //Remove any worksheet on second reference of the address. 
                    }
                }
                return;
            }
        }
        pos = this._ws.IndexOf('!');

        if (pos==0)
        {
            this._address = this._ws.Substring(1);
            this._ws = "";
            //_wb = "";
        }
        else if (pos > -1)
        {
            this._address = this._ws.Substring(pos + 1);
            this._ws = this._ws.Substring(0, pos);
        }
        else
        {
            this._address = address;
        }
        if(string.IsNullOrEmpty(this._address))
        {
            this._address = "A:XFD";
        }
    }
    internal void ChangeWorksheet(string wsName, string newWs)
    {
        if (this._ws == wsName)
        {
            this._ws = newWs;
        }

        string? fullAddress = this.GetAddress();
            
        if (this.Addresses != null)
        {
            foreach (ExcelAddressBase? a in this.Addresses)
            {
                if (a._ws == wsName)
                {
                    a._ws = newWs;
                    fullAddress += "," + a.GetAddress();
                }
                else
                {
                    fullAddress += "," + a._address;
                }
            }
        }

        this._address = fullAddress;
    }

    private string GetAddress()
    {
        string address = this.GetAddressWorkBookWorkSheet();
        if (this.IsName)
        {
            return address + GetAddress(this._fromRow, this._fromCol, this._toRow, this._toCol);
        }
        else
        {
            return address + GetAddress(this._fromRow, this._fromCol, this._toRow, this._toCol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed);
        }
    }

    internal string GetAddressWorkBookWorkSheet()
    {
        string? address = "";

        if (string.IsNullOrEmpty(this._ws) == false)
        {
            if (string.IsNullOrEmpty(this._wb) == false)
            {
                address = "[" + this._wb + "]";
            }

            if (this._address.IndexOf("'!", StringComparison.OrdinalIgnoreCase) >=0 || ExcelWorksheet.NameNeedsApostrophes(this._ws))
            {
                address += string.Format("'{0}'!", this._ws.Replace("'","''"));
            }
            else
            {
                address += string.Format("{0}!", this._ws);
            }
        }

        return address;
    }
    #endregion
    internal ExcelCellAddress _start = null;
    /// <summary>
    /// Gets the row and column of the top left cell.
    /// </summary>
    /// <value>The start row column.</value>
    public ExcelCellAddress Start
    {
        get { return this._start ??= new ExcelCellAddress(this._fromRow, this._fromCol, this._fromRowFixed, this._fromColFixed); }
    }
    internal ExcelCellAddress _end = null;
    /// <summary>
    /// Gets the row and column of the bottom right cell.
    /// </summary>
    /// <value>The end row column.</value>
    public ExcelCellAddress End
    {
        get { return this._end ??= new ExcelCellAddress(this._toRow, this._toCol, this._toRowFixed, this._toColFixed); }
    }
    /// <summary>
    /// The index to the external reference. Return 0, the current workbook, if no reference exists.
    /// </summary>
    public int ExternalReferenceIndex
    {
        get
        {
            if(this.Address.StartsWith("["))
            {
                if(this._wb.Any(x=>char.IsDigit(x)))
                {
                    return int.Parse(this._wb);
                }
                else
                {
                    return -1;
                }
            }
            else
            {
                return 0;
            }
        }
    }
    internal ExcelTableAddress _table = null;
    /// <summary>
    /// If the address is refering a table, this property contains additional information 
    /// </summary>
    public ExcelTableAddress Table
    {
        get
        {
            return this._table;
        }
    }

    /// <summary>
    /// The address for the range
    /// </summary>
    public virtual string Address
    {
        get
        {
            return this._address;
        }
    }
    /// <summary>
    /// The full address including the worksheet
    /// </summary>
    public string FullAddress
    {
        get
        {
            string a="";
            if(this._addresses != null)
            {
                foreach(ExcelAddressBase? sa in this._addresses)
                {
                    a += ","+sa.GetAddress();
                }
                a = a.TrimStart(',');
            }
            else
            {
                a = this.GetAddress();
            }
            return a;
        }
    }
    /// <summary>
    /// If the address is a defined name
    /// </summary>
    public bool IsName
    {
        get
        {
            return this._fromRow < 0;
        }
    }
    /// <summary>
    /// Returns the address text
    /// </summary>
    /// <returns></returns>
    public override string ToString()
    {
        return this._address;
    }
    /// <summary>
    /// Serves as the default hash function.
    /// </summary>
    /// <returns>A hash code for the current object.</returns>
    public override int GetHashCode()
    {
        return base.GetHashCode();
    }
    string _firstAddress;
    /// <summary>
    /// returns the first address if the address is a multi address.
    /// A1:A2,B1:B2 returns A1:A2
    /// </summary>
    internal string FirstAddress
    {
        get
        {
            if (string.IsNullOrEmpty(this._firstAddress))
            {
                return this._address;
            }
            else
            {
                return this._firstAddress;
            }
        }
    }
    /// <summary>
    /// Returns the address of the first cell in the address without $. Returns #REF! if the address is invalid.
    /// </summary>
    internal string FirstCellAddressRelative
    {
        get
        {
            if (this._fromRow > 0 && this._fromCol > 0)
            {
                return GetAddress(this._fromRow, this._fromCol);
            }
            return "#REF!";
        }
    }
    internal string AddressSpaceSeparated
    {
        get
        {
            return this._address.Replace(',', ' '); //Conditional formatting and a few other places use space as separator for mulit addresses.
        }
    }
    /// <summary>
    /// Validate the address
    /// </summary>
    protected void Validate()
    {
        if ((this._fromRow > this._toRow || this._fromCol > this._toCol) && this._toRow!=0) //_toRow==0 is #REF!
        {
            throw new ArgumentOutOfRangeException("Start cell Address must be less or equal to End cell address");
        }
    }
    internal string WorkSheetName
    {
        get
        {
            return this._ws;
        }
    }
    internal List<ExcelAddressBase> _addresses = null;
    internal virtual List<ExcelAddressBase> Addresses
    {
        get
        {
            return this._addresses;
        }
    }
    internal virtual List<ExcelAddressBase> GetAllAddresses()
    {
        if(this.Addresses==null)
        {
            return new List<ExcelAddressBase>() { this };
        }
        return this._addresses;
    }

    private bool ExtractAddress(string fullAddress)
    {
        Stack<int>? brackPos=new Stack<int>();
        List<string>? bracketParts=new List<string>();
        string first="", second="";
        bool isText=false, hasSheet=false, hasColon=false;
        string ws="";
        this._addresses = null;            
        try
        {
            if (fullAddress == "#REF!")
            {
                this.SetAddress(ref fullAddress, ref second, ref hasSheet);
                return true;
            }
            else if (Utils.ConvertUtil._invariantCompareInfo.IsPrefix(fullAddress, "!"))
            {
                // invalid address!
                return false;
            }
            for (int i = 0; i < fullAddress.Length; i++)
            {
                char c = fullAddress[i];
                if (c == '\'')
                {
                    if (isText && i + 1 < fullAddress.Length && fullAddress[i + 1] == '\'')
                    {
                        if (hasColon)
                        {
                            second += c;
                        }
                        else
                        {
                            first += c;
                        }
                    }
                    isText = !isText;
                }
                else
                {
                    if (brackPos.Count > 0)
                    {
                        if (c == '[' && !isText)
                        {
                            brackPos.Push(i);
                        }
                        else if (c == ']' && !isText)
                        {
                            if (brackPos.Count > 0)
                            {
                                int from = brackPos.Pop();
                                bracketParts.Add(fullAddress.Substring(from + 1, i - from - 1));

                                if (brackPos.Count == 0)
                                {
                                    this.HandleBrackets(first, second, bracketParts);
                                }
                            }
                            else
                            {
                                //Invalid address!
                                return false;
                            }
                        }
                    }
                    else if (c == ':' && !isText)
                    {
                        hasColon = true;
                    }
                    else if (c == '[' && !isText)
                    {
                        brackPos.Push(i);
                    }
                    else if (c == '!' && !isText && !first.EndsWith("#REF") && !second.EndsWith("#REF"))
                    {
                        // the following is to handle addresses that specifies the
                        // same worksheet twice: Sheet1!A1:Sheet1:A3
                        // They will be converted to: Sheet1!A1:A3
                        if (hasSheet && second != null && second.ToLower().EndsWith(first.ToLower()))
                        {
                            second = Regex.Replace(second, $"{first}$", string.Empty);
                        }
                        if (string.IsNullOrEmpty(ws))
                        {                                
                            if (second == "")
                            {
                                ws = first;
                                first = "";
                            }
                            else
                            {
                                ws = second;
                                second = "";
                            }
                        }
                        else if(string.IsNullOrEmpty(second)==false)
                        {
                            if(!ws.Equals(second,StringComparison.OrdinalIgnoreCase))
                            {
                                this._fromRow = this._toRow = this._fromCol = this._toCol = -1;
                                return true;
                            }
                            second = "";
                        }
                        hasSheet = true;
                    }
                    else if (c == ',' && !isText)
                    {
                        this._addresses ??= new List<ExcelAddressBase>();

                        if(string.IsNullOrEmpty(ws))
                        {
                            first = string.IsNullOrEmpty(second) ? first : first + ":" + second;
                            second = "";
                        }
                        else
                        {
                            second = string.IsNullOrEmpty(second) ? first : first + ":" + second;
                            first = ws;
                        }

                        this.SetAddress(ref first, ref second, ref hasSheet);
                        ws = "";
                        hasSheet = false;
                        hasColon = false;
                    }
                    else
                    {
                        if (hasColon)
                        {
                            second += c;
                        }
                        else
                        {
                            first += c;
                        }
                    }
                }
            }
            if (this.Table == null)
            {
                if (string.IsNullOrEmpty(ws))
                {
                    first = string.IsNullOrEmpty(second) ? first : first + ":" + second;
                    second = "";
                }
                else
                {
                    second = string.IsNullOrEmpty(second) ? first : first + ":" + second;
                    first = ws;
                }

                this.SetAddress(ref first, ref second, ref hasSheet);
            }
            return true;
        }
        catch
        {
            return false;
        }
    }

    private void HandleBrackets(string first, string second, List<string> bracketParts)
    {
        if(!string.IsNullOrEmpty(first))
        {
            this._table = new ExcelTableAddress();
            this.Table.Name = first;
            foreach (string? s in bracketParts)
            {
                if(s.IndexOf('[')<0)
                {
                    switch(s.ToLower(CultureInfo.InvariantCulture))                
                    {
                        case "#all":
                            this._table.IsAll = true;
                            break;
                        case "#headers":
                            this._table.IsHeader = true;
                            break;
                        case "#data":
                            this._table.IsData = true;
                            break;
                        case "#totals":
                            this._table.IsTotals = true;
                            break;
                        case "#this row":
                            this._table.IsThisRow = true;
                            break;
                        default:
                            if(string.IsNullOrEmpty(this._table.ColumnSpan))
                            {
                                this._table.ColumnSpan=s;
                            }
                            else
                            {
                                this._table.ColumnSpan += ":" + s;
                            }
                            break;
                    }                
                }
            }
        }
    }
    #region Address manipulation methods
    internal eAddressCollition Collide(ExcelAddressBase address, bool ignoreWs=false)
    {
        if (ignoreWs == false && address.WorkSheetName != this.WorkSheetName && 
            string.IsNullOrEmpty(address.WorkSheetName) == false && 
            string.IsNullOrEmpty(this.WorkSheetName) == false)
        {
            return eAddressCollition.No;
        }

        return this.Collide(address._fromRow, address._fromCol, address._toRow, address._toCol);
    }
    internal eAddressCollition Collide(int row, int col)
    {
        return this.Collide(row, col, row, col);
    }
    internal eAddressCollition Collide(int fromRow, int fromCol, int toRow, int toCol)
    {
        if (this.DoNotCollide(fromRow, fromCol, toRow, toCol))
        {
            return eAddressCollition.No;
        }
        else if (fromRow == this._fromRow && fromCol == this._fromCol &&
                 toRow == this._toRow && toCol == this._toCol)
        {
            return eAddressCollition.Equal;
        }
        else if (fromRow >= this._fromRow && toRow <= this._toRow &&
                 fromCol >= this._fromCol && toCol <= this._toCol)
        {
            return eAddressCollition.Inside;
        }
        else
        {
            return eAddressCollition.Partly;
        }
    }

    internal bool DoNotCollide(int fromRow, int fromCol, int toRow, int toCol)
    {
        return fromRow > this._toRow || fromCol > this._toCol
                                     || this._fromRow > toRow || this._fromCol > toCol;
    }

    internal bool CollideFullRowOrColumn(ExcelAddressBase address)
    {
        return this.CollideFullRowOrColumn(address._fromRow, address._fromCol, address._toRow, address._toCol);
    }
    internal bool CollideFullRowOrColumn(int fromRow, int fromCol, int toRow, int toCol)
    {
        return (this.CollideFullRow(fromRow, toRow) && this.CollideColumn(fromCol, toCol)) || 
               (this.CollideFullColumn(fromCol, toCol) && this.CollideRow(fromRow, toRow));
    }
    private bool CollideColumn(int fromCol, int toCol)
    {
        return fromCol  <= this._toCol && toCol >= this._fromCol;
    }

    internal bool CollideRow(int fromRow, int toRow)
    {
        return fromRow <= this._toRow && toRow >= this._fromRow;
    }
    internal bool CollideFullRow(int fromRow, int toRow)
    {
        return fromRow <= this._fromRow && toRow >= this._toRow;
    }
    internal bool CollideFullColumn(int fromCol, int toCol)
    {
        return fromCol <= this._fromCol && toCol >= this._toCol;
    }
    internal ExcelAddressBase AddRow(int row, int rows, bool setFixed=false, bool setRefOnMinMax=true, bool extendIfLastRow=false)
    {
        if (row > this._toRow && (row!= this._toRow+1 || extendIfLastRow==false))
        {
            return this;
        }
        int toRow = setFixed && this._toRowFixed ? this._toRow : this._toRow + rows;
        if (toRow < 1)
        {
            return null;
        }

        if (row <= this._fromRow)
        {
            int fromRow = setFixed && this._fromRowFixed ? this._fromRow : this._fromRow + rows;
            if (fromRow > ExcelPackage.MaxRows)
            {
                return null;
            }

            return new ExcelAddressBase(GetRow(fromRow, setRefOnMinMax), this._fromCol, GetRow(toRow, setRefOnMinMax), this._toCol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed, this.WorkSheetName, this._address);
        }
        else
        {
            return new ExcelAddressBase(this._fromRow, this._fromCol, GetRow(toRow, setRefOnMinMax), this._toCol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed, this.WorkSheetName, this._address);
        }
    }

    private static int GetRow(int row, bool setRefOnMinMax)
    {
        if (setRefOnMinMax==false)
        {
            if (row < 1)
            {
                return 1;
            }

            if (row > ExcelPackage.MaxRows)
            {
                return ExcelPackage.MaxRows;
            }
        }

        return row;
    }
    private static int GetColumn(int column, bool setRefOnMinMax)
    {
        if (setRefOnMinMax == false)
        {
            if (column < 1)
            {
                return 1;
            }

            if (column > ExcelPackage.MaxColumns)
            {
                return ExcelPackage.MaxColumns;
            }
        }

        return column;
    }

    internal ExcelAddressBase DeleteRow(int row, int rows, bool setFixed = false, bool adjustMaxRow=true)
    {
        if (row > this._toRow) //After
        {
            return this;
        }
        else if (row != 0 && row <= this._fromRow && row + rows > this._toRow) //Inside
        {
            return null;
        }
        else if (row+rows < this._fromRow || (this._fromRowFixed && row < this._fromRow)) //Before
        {
            int toRow = (setFixed && this._toRowFixed) || (adjustMaxRow==false && this._toRow==ExcelPackage.MaxRows) ? this._toRow : this._toRow - rows;
            return new ExcelAddressBase(setFixed && this._fromRowFixed ? this._fromRow : this._fromRow - rows, this._fromCol, toRow, this._toCol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed, this.WorkSheetName, this._address);
        }
        else  //Partly
        {
            if (row <= this._fromRow)
            {
                int toRow = (setFixed && this._toRowFixed) || (adjustMaxRow == false && this._toRow == ExcelPackage.MaxRows) ? this._toRow : this._toRow - rows;

                return new ExcelAddressBase(row, this._fromCol, toRow, this._toCol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed, this.WorkSheetName, this._address);
            }
            else
            {
                int toRow = (setFixed && this._toRowFixed) || (adjustMaxRow == false && this._toRow == ExcelPackage.MaxRows) ? this._toRow : this._toRow - rows < row ? row - 1 : this._toRow - rows;
                return new ExcelAddressBase(this._fromRow, this._fromCol, toRow, this._toCol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed, this.WorkSheetName, this._address);
            }
        }
    }
    internal ExcelAddressBase AddColumn(int col, int cols, bool setFixed = false, bool setRefOnMinMax=true)
    {
        if (col > this._toCol)
        {
            return this;
        }
        int toCol = GetColumn(setFixed && this._toColFixed ? this._toCol : this._toCol + cols, setRefOnMinMax);
        if (col <= this._fromCol)
        {
            int fromCol = GetColumn(setFixed && this._fromColFixed ? this._fromCol : this._fromCol + cols, setRefOnMinMax);
            return new ExcelAddressBase(this._fromRow, fromCol, this._toRow, toCol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed, this.WorkSheetName, this._address);
        }
        else
        {
            return new ExcelAddressBase(this._fromRow, this._fromCol, this._toRow, toCol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed, this.WorkSheetName, this._address);
        }
    }
    internal ExcelAddressBase DeleteColumn(int col, int cols, bool setFixed = false, bool adjustMaxCol = true)
    {
        if (col > this._toCol) //After
        {
            return this;
        }
        if (col!=0 && col <= this._fromCol && col + cols > this._toCol) //Inside
        {
            return null;
        }
        else if (col + cols < this._fromCol || (this._fromColFixed && col < this._fromCol)) //Before
        {
            int toCol = (setFixed && this._toColFixed) ||(adjustMaxCol==false && this._toCol==ExcelPackage.MaxColumns) ? this._toCol : this._toCol - cols;
            return new ExcelAddressBase(this._fromRow, setFixed && this._fromColFixed ? this._fromCol : this._fromCol - cols, this._toRow, toCol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed, this.WorkSheetName, this._address);
        }
        else  //Partly
        {
            if (col <= this._fromCol)
            {
                int toCol = (setFixed && this._toColFixed) || (adjustMaxCol == false && this._toCol == ExcelPackage.MaxColumns) ? this._toCol : this._toCol - cols;
                return new ExcelAddressBase(this._fromRow, col, this._toRow, toCol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed, this._ws, this._address);
            }
            else
            {
                int toCol = (setFixed && this._toColFixed) || (adjustMaxCol == false && this._toCol == ExcelPackage.MaxColumns) ? this._toCol : this._toCol - cols < col ? col - 1 : this._toCol - cols;
                return new ExcelAddressBase(this._fromRow, this._fromCol, this._toRow, toCol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed, this._ws, this._address);
            }
        }
    }
    internal ExcelAddressBase Insert(ExcelAddressBase address, eShiftTypeInsert Shift)
    {
        //Before or after, no change
        if(this._toRow < address._fromRow || this._toCol < address._fromCol || (this._fromRow > address._toRow && this._fromCol > address._toCol))
        {
            return this;
        }

        int rows = address.Rows;
        int cols = address.Columns;
        string retAddress = "";
        if (Shift==eShiftTypeInsert.Right)
        {
            if (address._fromRow > this._fromRow)
            {
                retAddress = GetAddress(this._fromRow, this._fromCol, address._fromRow, this._toCol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed);
            }
            if(address._fromCol > this._fromCol)
            {
                retAddress = GetAddress(this._fromRow < address._fromRow ? this._fromRow : address._fromRow, this._fromCol, address._fromRow, this._toCol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed);
            }
        }
        if (this._toRow < address._fromRow)
        {
            if (this._fromRow < address._fromRow)
            {

            }
            else
            {
            }
        }
        return null;
    }
    #endregion
    private void SetAddress(ref string first, ref string second, ref bool hasSheet)
    {
        string ws, address;
        if (hasSheet)
        {
            ws = first;
            address = second;
            first = "";
            second = "";
        }
        else
        {
            address = first;
            ws = "";
            first = "";
        }
        hasSheet = false;
        if (string.IsNullOrEmpty(this._firstAddress))
        {
            if (string.IsNullOrEmpty(this._ws) || !string.IsNullOrEmpty(ws))
            {
                this._ws = ws;                    
            }

            this._firstAddress = address;
            GetRowColFromAddress(address, out this._fromRow, out this._fromCol, out this._toRow, out this._toCol, out this._fromRowFixed, out this._fromColFixed, out this._toRowFixed, out this._toColFixed);
            this._start = null;
            this._end = null;
        }
        if (this._addresses != null)
        {
            this._addresses.Add(new ExcelAddress(this._ws, address));
        }
    }
    internal enum AddressType
    {
        Invalid,
        InternalAddress,
        ExternalAddress,
        InternalName,
        ExternalName,
        Formula,
        R1C1
    }

    internal static AddressType IsValid(string Address, bool r1c1=false)
    {
        double d;
        if (Address == "#REF!")
        {
            return AddressType.Invalid;
        }
        else if(double.TryParse(Address, NumberStyles.Any, CultureInfo.InvariantCulture, out d)) //A double, no valid address
        {
            return AddressType.Invalid;
        }
        else if (IsFormula(Address))
        {
            return AddressType.Formula;
        }
        else
        {
            if (r1c1 && IsR1C1(Address))
            {
                return AddressType.R1C1;
            }
            else
            {
                string ws;
                if (SplitAddress(Address, out string wb, out ws, out string intAddress))
                {

                    if (intAddress.Contains("[")) //Table reference
                    {
                        return string.IsNullOrEmpty(wb) ? AddressType.InternalAddress : AddressType.ExternalAddress;
                    }
                    if (intAddress.Contains(","))
                    {
                        intAddress = intAddress.Substring(0, intAddress.IndexOf(','));
                    }
                    if (IsAddress(intAddress, true))
                    {
                        return string.IsNullOrEmpty(wb) ? AddressType.InternalAddress : AddressType.ExternalAddress;
                    }
                    else
                    {
                        return string.IsNullOrEmpty(wb) ? AddressType.InternalName : AddressType.ExternalName;
                    }
                }
                else
                {
                    return AddressType.Invalid;
                }
            }
        }
    }
    private static bool IsR1C1(string address)
    {
        int start = address.LastIndexOf("!", address.Length-1, StringComparison.OrdinalIgnoreCase);
        if (start>=0)
        {
            address = address.Substring(start + 1);
        }
        address = address.ToUpper();
        if (string.IsNullOrEmpty(address) || (address[0]!='R' && address[0]!='C'))
        {
            return false;
        }
        bool isC = false, isROrC = false;
        bool startBracket = false;
        foreach(char c in address)
        {
            switch(c)
            {
                case 'C':
                    isC = true;
                    isROrC = true;
                    break;
                case 'R':
                    if (isC)
                    {
                        return false;
                    }

                    isROrC = true;
                    break;
                case '[':
                    startBracket = true;
                    break;
                case ']':
                    if (startBracket == false)
                    {
                        return false;
                    }

                    isROrC = false;
                    break;
                case ':':
                    isC = false;
                    startBracket = false;
                    isROrC = false;
                    break;
                default:
                    if((c>='0' && c<='9') ||c=='-')
                    {
                        if(isROrC==false)
                        {
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                    break;
            }
        }
        return true;
    }

    private static bool IsAddress(string intAddress, bool allowRef = false)
    {
        if(string.IsNullOrEmpty(intAddress))
        {
            return false;
        }

        string[]? cells = intAddress.Split(':');
        int toRow,
            toCol;

        if(!GetRowCol(cells[0], out int fromRow, out int fromCol, false))
        {
            return false;
        }
        if (cells.Length > 1)
        {
            if (!GetRowCol(cells[1], out toRow, out toCol, false))
            {
                return false;
            }
        }
        else
        {
            toRow = fromRow;
            toCol = fromCol;
        }
        if (allowRef)
        {
            return
                fromCol > -1 &&
                toCol <= ExcelPackage.MaxColumns &&
                fromRow > -1 &&
                toRow <= ExcelPackage.MaxRows;
        }
        else
        {
            return 
                fromRow <= toRow &&
                fromCol <= toCol &&
                fromCol > -1 &&
                toCol <= ExcelPackage.MaxColumns &&
                fromRow > -1 &&
                toRow <= ExcelPackage.MaxRows;
        }
    }

    private static bool SplitAddress(string Address, out string wb, out string ws, out string intAddress)
    {
        wb = "";
        ws = "";
        intAddress = "";
        string? text = "";
        bool isText = false;
        int brackPos=-1;
        for (int i = 0; i < Address.Length; i++)
        {
            if (Address[i] == '\'')
            {
                isText = !isText;
                if(i>0 && Address[i-1]=='\'')
                {
                    text += "'";
                }
            }
            else
            {
                if(Address[i]=='!' && !isText)
                {
                    if (text.Length>0 && text[0] == '[')
                    {
                        wb = text.Substring(1, text.IndexOf(']') - 1);
                        ws = text.Substring(text.IndexOf(']') + 1);
                    }
                    else
                    {
                        ws=text;
                    }
                    intAddress=Address.Substring(i+1);
                    return true;
                }
                else
                {
                    if(Address[i]=='[' && !isText)
                    {
                        if (i > 0) //Table reference return full address;
                        {
                            intAddress=Address;
                            return true;
                        }
                        brackPos=i;
                    }
                    else if(Address[i]==']' && !isText)
                    {
                        if (brackPos > -1)
                        {
                            wb = text;
                            text = "";
                        }
                        else
                        {
                            return false;
                        }
                    }
                    else
                    {
                        text+=Address[i];
                    }
                }
            }
        }
        intAddress = text;
        return true;
    }

    private static readonly HashSet<char> _tokens = new HashSet<char>(new char[] { '+', '-', '*', '/', '^', '&', '=', '<', '>', '(', ')', '{', '}', '%', '\"' }); //See TokenSeparatorProvider
    internal static bool IsFormula(string address)
    {
        bool isText = false;
        int tableNameCount = 0;
        for (int i = 0; i < address.Length; i++)
        {
            char addressChar = address[i];
            if (addressChar == '\'')
            {
                if(i>0 && isText==false && address.Length>i+1 && address[i - 1] == ' ' && address[i+1] != '\'')
                {
                    return true;
                }
                isText = !isText;
            }
            else if (isText == false && addressChar == '[')
            {
                tableNameCount++;
            }
            else if (isText == false && addressChar == ']')
            {
                tableNameCount--;
            }
            else if(tableNameCount==0)
            {
                if (isText == false && _tokens.Contains(addressChar))
                {
                    return true;
                }
            }
        }
        return false;
    }
    private static bool IsValidName(string address)
    {
        if (Regex.IsMatch(address, "[^0-9./*-+,½!\"@#£%&/{}()\\[\\]=?`^~':;<>|][^/*-+,½!\"@#£%&/{}()\\[\\]=?`^~':;<>|]*"))
        {
            return true;
        }
        else
        {
            return false;
        }
    }

    /// <summary>
    /// Number of rows int the address
    /// </summary>
    public int Rows 
    {
        get
        {
            return this._toRow - this._fromRow+1;
        }
    }
    /// <summary>
    /// Number of columns int the address
    /// </summary>
    public int Columns
    {
        get
        {
            return this._toCol - this._fromCol + 1;
        }
    }
    /// <summary>
    /// Returns true if the range spans a full row
    /// </summary>
    /// <returns></returns>
    public bool IsFullRow
    {
        get
        {
            return this._fromCol == 1 && this._toCol == ExcelPackage.MaxColumns;
        }
    }
    /// <summary>
    /// Returns true if the range spans a full column
    /// </summary>
    /// <returns></returns>
    public bool IsFullColumn
    {
        get
        {
            return this._fromRow == 1 && this._toRow == ExcelPackage.MaxRows;
        }
    }

    internal bool IsSingleCell
    {
        get
        {
            return this._fromRow == this._toRow && this._fromCol == this._toCol;
        }
    }

    /// <summary>
    /// The address without the workbook or worksheet reference
    /// </summary>
    public string LocalAddress 
    { 
        get
        {                
            if (this.Addresses == null)
            {
                if (this._table == null)
                {
                    return GetAddress(this._fromRow, this._fromCol, this._toRow, this._toCol, this._fromRowFixed, this._fromColFixed, this._toRowFixed, this._toColFixed);
                }
                else
                {
                    return RemoveSheetName(this.FirstAddress);
                }
            }
            else
            {
                StringBuilder? sb = new StringBuilder();
                foreach (ExcelAddressBase? a in this.Addresses)
                {
                    if (a._table == null)
                    {
                        sb.Append(GetAddress(a._fromRow, a._fromCol, a._toRow, a._toCol, a._fromRowFixed, a._fromColFixed, a._toRowFixed, a._toColFixed));
                    }
                    else
                    {
                        sb.Append(RemoveSheetName(a.Address));
                    }
                    sb.Append(",");
                }
                return sb.ToString(0, sb.Length - 1);
            }
        }
    }

    private static string RemoveSheetName(string address)
    {
        int ix = address.TrimEnd().LastIndexOf('!', address.Length - 2);  //Last index can be ! if address is #REF!, so check from 
        if (ix >= 0)
        {
            address = address.Substring(ix + 1);
        }

        return address;
    }

    /// <summary>
    /// The address without the workbook reference
    /// </summary>
    internal string WorkbookLocalAddress
    {
        get
        {
            if (!this._address.StartsWith("["))
            {
                return this._address;
            }

            int ix = this._address.IndexOf("]",1);
            if (ix >= 0)
            {
                return this._address.Substring(ix + 1);
            }
            return this._address;
        }
    }

    internal static string GetWorkbookPart(string address)
    {
        int ix = 0;
        if(address[ix]=='\'')
        {
            ix++;
        }
        if (address[ix] == '[')
        {
            int endIx = address.LastIndexOf(']');
            if (endIx > 0)
            {
                return address.Substring(ix+1, endIx - ix - 1);
            }   
        }
        return "";
    }
    internal static string GetWorksheetPart(string address, string defaultWorkSheet)
    {
        int ix=0;
        return GetWorksheetPart(address, defaultWorkSheet, ref ix);
    }
    internal static string GetWorksheetPart(string address, string defaultWorkSheet, ref int endIx)
    {
        if(address=="")
        {
            return defaultWorkSheet;
        }

        int ix = 0;
        if (address[0] == '[' || address.StartsWith("'["))
        {
            ix = address.IndexOf(']')+1;
        }
        if (ix >= 0 && ix < address.Length)
        {
            if (address[ix] == '\'')
            {
                string? ret=GetString(address, ix+1, out endIx);
                endIx++;
                return ret; 
            }
            else
            {
                endIx = address.IndexOf('!',ix)+1;
                int subtrLen = 1;
                if(endIx>0 && address[endIx-2]=='\'')
                {
                    subtrLen++;
                }
                if(endIx > ix)
                {
                    return address.Substring(ix, endIx - ix - subtrLen);
                }   
                else
                {
                    return defaultWorkSheet;
                }
            }
        }
        else
        {
            return defaultWorkSheet;
        }
    }
    internal static string GetAddressPart(string address)
    {
        int ix=0;
        GetWorksheetPart(address, "", ref ix);
        if(ix<address.Length)
        {
            if (address[ix] == '!')
            {
                return address.Substring(ix + 1);
            }
            else
            {
                return "";
            }
        }
        else
        {
            return "";
        }

    }
    internal static void SplitAddress(string fullAddress, out string wb, out string ws, out string address, string defaultWorksheet="")
    {
        wb = GetWorkbookPart(fullAddress);
        int ix=0;
        ws = GetWorksheetPart(fullAddress, defaultWorksheet, ref ix);
        if (ix < fullAddress.Length)
        {
            if (fullAddress[ix] == '!')
            {
                address = fullAddress.Substring(ix + 1);
            }
            else
            {
                address = fullAddress.Substring(ix);
            }
        }
        else
        {
            address="";
        }
    }
    internal static List<string[]> SplitFullAddress(string fullAddress)
    {
        List<string[]>? addresses = new List<string[]>();
        string[]? currentAddress = new string[3];
        bool isInWorkbook = false;
        bool isInWorksheet = false;
        bool isInAddress = false;
        bool isInText = false;
        int prevPos = 0;
        for (int i=0;i<fullAddress.Length;i++)
        {
            if (isInWorkbook == false &&
                isInWorksheet == false &&
                isInAddress == false)
            {
                if (fullAddress[i] == '[')
                {
                    isInWorkbook = true;
                    prevPos = i + 1;
                }
                else if(fullAddress[i] == '\'')
                {
                    isInWorksheet = true;
                    isInText = true;
                    prevPos = i + 1;
                }
                else if (fullAddress[i]=='!')
                {
                    isInAddress = true;
                    prevPos = i + 1;
                }
                else
                {
                    isInAddress = true;
                }
            }
            else if(isInWorkbook)
            {
                if (fullAddress[i] == ']')
                {
                    currentAddress[0] = fullAddress.Substring(prevPos, i - prevPos);
                    isInWorkbook = false;
                }
            }
            else if(isInWorksheet)
            {
                if (fullAddress[i] == '\'')
                {
                    isInText = !isInText;
                }
                else if (isInText==false && fullAddress[i] == '!')
                {
                    currentAddress[1] = fullAddress.Substring(prevPos, i -prevPos - 1).Replace("''","'");
                    prevPos = i + 1;
                    isInWorksheet = false;
                }
            }
            else if(isInAddress)
            {
                if(fullAddress[i] == '!')
                {
                    currentAddress[1] = fullAddress.Substring(prevPos, i - prevPos);
                    prevPos = i + 1;
                }
                else if (fullAddress[i]==',')
                {
                    currentAddress[2] = fullAddress.Substring(prevPos, i - prevPos);
                    addresses.Add(currentAddress);
                    prevPos = i + 1;
                    isInAddress = false;
                    currentAddress = new string[3];
                }
            }
        }

        if(isInWorkbook || isInWorksheet)
        {
            throw new ArgumentException($"Invalid address {fullAddress}");
        }
        currentAddress[2] = fullAddress.Substring(prevPos, fullAddress.Length - prevPos);
        addresses.Add(currentAddress);
        return addresses;
    }
    private static string GetString(string address, int ix, out int endIx)
    {
        int strIx = address.IndexOf("''", ix);
        int prevStrIx = ix;
        while (strIx > -1)
        {
            prevStrIx = strIx;
            strIx = address.IndexOf("''", strIx + 1);
        }
        endIx = address.IndexOf("'", prevStrIx + 1) + 1;
        return address.Substring(ix, endIx - ix - 1).Replace("''", "'");
    }

    internal bool IsValidRowCol()
    {
        return !(this._fromRow > this._toRow  || this._fromCol > this._toCol || this._fromRow < 1 || this._fromCol < 1 || this._toRow > ExcelPackage.MaxRows || this._toCol > ExcelPackage.MaxColumns);
    }
    /// <summary>
    /// Returns true if the item is equal to another item.
    /// </summary>
    /// <param name="obj">The item to compare</param>
    /// <returns>True if the items are equal</returns>
    public override bool Equals(object obj)
    {
        if (obj is ExcelAddressBase a)
        {
            if (this.Addresses==null || a.Addresses==null)
            {
                if (this.Addresses?.Count > 1 || a.Addresses?.Count > 1)
                {
                    return false;
                }

                return IsEqual(this, a);
            }
            else
            {
                if (this.Addresses.Count != a.Addresses.Count)
                {
                    return false;
                }

                for(int i=0;i< this.Addresses.Count;i++)
                {
                    if (IsEqual(this.Addresses[i], a.Addresses[i]) == false)
                    {
                        return false;
                    }
                }
                return true;
            }
        }
        else
        {
            return this._address == obj?.ToString();
        }
    }

    private static bool IsEqual(ExcelAddressBase a1, ExcelAddressBase a2)
    {
        return a1._fromRow == a2._fromRow &&
               a1._toRow == a2._toRow &&
               a1._fromCol == a2._fromCol &&
               a1._toCol == a2._toCol;
    }
    /// <summary>
    /// Returns true the address contains an external reference
    /// </summary>
    public bool IsExternal
    {
        get
        {
            return !string.IsNullOrEmpty(this._wb);
        }
    }
}