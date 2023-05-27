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
using System.Text;
using System.Xml;
using System.Globalization;
using System.Linq;

namespace OfficeOpenXml;

#region "Enums"

/// <summary>
/// Printer orientation
/// </summary>
public enum eOrientation
{
    /// <summary>
    /// Portrait orientation
    /// </summary>
    Portrait,

    /// <summary>
    /// Landscape orientation
    /// </summary>
    Landscape
}

/// <summary>
/// Papersize
/// </summary>
public enum ePaperSize
{
    /// <summary>
    /// Letter paper (8.5 in. by 11 in.)
    /// </summary>
    Letter = 1,

    /// <summary>
    /// Letter small paper (8.5 in. by 11 in.)
    /// </summary>
    LetterSmall = 2,

    /// <summary>
    /// // Tabloid paper (11 in. by 17 in.)
    /// </summary>
    Tabloid = 3,

    /// <summary>
    /// Ledger paper (17 in. by 11 in.)
    /// </summary>
    Ledger = 4,

    /// <summary>
    /// Legal paper (8.5 in. by 14 in.)
    /// </summary>
    Legal = 5,

    /// <summary>
    /// Statement paper (5.5 in. by 8.5 in.)
    /// </summary>
    Statement = 6,

    /// <summary>
    /// Executive paper (7.25 in. by 10.5 in.)
    /// </summary>
    Executive = 7,

    /// <summary>
    /// A3 paper (297 mm by 420 mm)
    /// </summary>
    A3 = 8,

    /// <summary>
    /// A4 paper (210 mm by 297 mm)
    /// </summary>
    A4 = 9,

    /// <summary>
    /// A4 small paper (210 mm by 297 mm)
    /// </summary>
    A4Small = 10,

    /// <summary>
    /// A5 paper (148 mm by 210 mm)
    /// </summary>
    A5 = 11,

    /// <summary>
    /// B4 paper (250 mm by 353 mm)
    /// </summary>
    B4 = 12,

    /// <summary>
    /// B5 paper (176 mm by 250 mm)
    /// </summary>
    B5 = 13,

    /// <summary>
    /// Folio paper (8.5 in. by 13 in.)
    /// </summary>
    Folio = 14,

    /// <summary>
    /// Quarto paper (215 mm by 275 mm)
    /// </summary>
    Quarto = 15,

    /// <summary>
    /// Standard paper (10 in. by 14 in.)
    /// </summary>
    Standard10_14 = 16,

    /// <summary>
    /// Standard paper (11 in. by 17 in.)
    /// </summary>
    Standard11_17 = 17,

    /// <summary>
    /// Note paper (8.5 in. by 11 in.)
    /// </summary>
    Note = 18,

    /// <summary>
    /// #9 envelope (3.875 in. by 8.875 in.)
    /// </summary>
    Envelope9 = 19,

    /// <summary>
    /// #10 envelope (4.125 in. by 9.5 in.)
    /// </summary>
    Envelope10 = 20,

    /// <summary>
    /// #11 envelope (4.5 in. by 10.375 in.)
    /// </summary>
    Envelope11 = 21,

    /// <summary>
    /// #12 envelope (4.75 in. by 11 in.)
    /// </summary>
    Envelope12 = 22,

    /// <summary>
    /// #14 envelope (5 in. by 11.5 in.)
    /// </summary>
    Envelope14 = 23,

    /// <summary>
    /// C paper (17 in. by 22 in.)
    /// </summary>
    C = 24,

    /// <summary>
    /// D paper (22 in. by 34 in.)
    /// </summary>
    D = 25,

    /// <summary>
    /// E paper (34 in. by 44 in.)
    /// </summary>
    E = 26,

    /// <summary>
    /// DL envelope (110 mm by 220 mm)
    /// </summary>
    DLEnvelope = 27,

    /// <summary>
    /// C5 envelope (162 mm by 229 mm)
    /// </summary>
    C5Envelope = 28,

    /// <summary>
    /// C3 envelope (324 mm by 458 mm)
    /// </summary>
    C3Envelope = 29,

    /// <summary>
    /// C4 envelope (229 mm by 324 mm)
    /// </summary>
    C4Envelope = 30,

    /// <summary>
    /// C6 envelope (114 mm by 162 mm)
    /// </summary>
    C6Envelope = 31,

    /// <summary>
    /// C65 envelope (114 mm by 229 mm)
    /// </summary>
    C65Envelope = 32,

    /// <summary>
    /// B4 envelope (250 mm by 353 mm)
    /// </summary>
    B4Envelope = 33,

    /// <summary>
    /// B5 envelope (176 mm by 250 mm)
    /// </summary>
    B5Envelope = 34,

    /// <summary>
    /// B6 envelope (176 mm by 125 mm)
    /// </summary>
    B6Envelope = 35,

    /// <summary>
    /// Italy envelope (110 mm by 230 mm)
    /// </summary>
    ItalyEnvelope = 36,

    /// <summary>
    /// Monarch envelope (3.875 in. by 7.5 in.).
    /// </summary>
    MonarchEnvelope = 37,

    /// <summary>
    /// 6 3/4 envelope (3.625 in. by 6.5 in.)
    /// </summary>
    Six3_4Envelope = 38,

    /// <summary>
    /// US standard fanfold (14.875 in. by 11 in.)
    /// </summary>
    USStandard = 39,

    /// <summary>
    /// German standard fanfold (8.5 in. by 12 in.)
    /// </summary>
    GermanStandard = 40,

    /// <summary>
    /// German legal fanfold (8.5 in. by 13 in.)
    /// </summary>
    GermanLegal = 41,

    /// <summary>
    /// ISO B4 (250 mm by 353 mm)
    /// </summary>
    ISOB4 = 42,

    /// <summary>
    ///  Japanese double postcard (200 mm by 148 mm)
    /// </summary>
    JapaneseDoublePostcard = 43,

    /// <summary>
    /// Standard paper (9 in. by 11 in.)
    /// </summary>
    Standard9 = 44,

    /// <summary>
    /// Standard paper (10 in. by 11 in.)
    /// </summary>
    Standard10 = 45,

    /// <summary>
    /// Standard paper (15 in. by 11 in.)
    /// </summary>
    Standard15 = 46,

    /// <summary>
    /// Invite envelope (220 mm by 220 mm)
    /// </summary>
    InviteEnvelope = 47,

    /// <summary>
    /// Letter extra paper (9.275 in. by 12 in.)
    /// </summary>
    LetterExtra = 50,

    /// <summary>
    /// Legal extra paper (9.275 in. by 15 in.)
    /// </summary>
    LegalExtra = 51,

    /// <summary>
    /// Tabloid extra paper (11.69 in. by 18 in.)
    /// </summary>
    TabloidExtra = 52,

    /// <summary>
    /// A4 extra paper (236 mm by 322 mm)
    /// </summary>
    A4Extra = 53,

    /// <summary>
    /// Letter transverse paper (8.275 in. by 11 in.)
    /// </summary>
    LetterTransverse = 54,

    /// <summary>
    /// A4 transverse paper (210 mm by 297 mm)
    /// </summary>
    A4Transverse = 55,

    /// <summary>
    /// Letter extra transverse paper (9.275 in. by 12 in.)
    /// </summary>
    LetterExtraTransverse = 56,

    /// <summary>
    /// SuperA/SuperA/A4 paper (227 mm by 356 mm)
    /// </summary>
    SuperA = 57,

    /// <summary>
    /// SuperB/SuperB/A3 paper (305 mm by 487 mm)
    /// </summary>
    SuperB = 58,

    /// <summary>
    /// Letter plus paper (8.5 in. by 12.69 in.)
    /// </summary>
    LetterPlus = 59,

    /// <summary>
    /// A4 plus paper (210 mm by 330 mm)
    /// </summary>
    A4Plus = 60,

    /// <summary>
    /// A5 transverse paper (148 mm by 210 mm)
    /// </summary>
    A5Transverse = 61,

    /// <summary>
    /// JIS B5 transverse paper (182 mm by 257 mm)
    /// </summary>
    JISB5Transverse = 62,

    /// <summary>
    /// A3 extra paper (322 mm by 445 mm)
    /// </summary>
    A3Extra = 63,

    /// <summary>
    /// A5 extra paper (174 mm by 235 mm)
    /// </summary>
    A5Extra = 64,

    /// <summary>
    /// ISO B5 extra paper (201 mm by 276 mm)
    /// </summary>
    ISOB5 = 65,

    /// <summary>
    /// A2 paper (420 mm by 594 mm)
    /// </summary>
    A2 = 66,

    /// <summary>
    /// A3 transverse paper (297 mm by 420 mm)
    /// </summary>
    A3Transverse = 67,

    /// <summary>
    /// A3 extra transverse paper (322 mm by 445 mm*/
    /// </summary>
    A3ExtraTransverse = 68
}

/// <summary>
/// Specifies printed page order
/// </summary>
public enum ePageOrder
{
    /// <summary>
    /// Order pages vertically first, then move horizontally.
    /// </summary>
    DownThenOver,

    /// <summary>
    /// Order pages horizontally first, then move vertically
    /// </summary>
    OverThenDown
}

#endregion

/// <summary>
/// Printer settings
/// </summary>
public sealed class ExcelPrinterSettings : XmlHelper
{
    ExcelWorksheet _ws;
    bool _marginsCreated = false;

    internal ExcelPrinterSettings(XmlNamespaceManager ns, XmlNode topNode, ExcelWorksheet ws)
        : base(ns, topNode)
    {
        this._ws = ws;
        this.SchemaNodeOrder = ws.SchemaNodeOrder;
    }

    const string _leftMarginPath = "d:pageMargins/@left";

    /// <summary>
    /// Left margin in inches
    /// </summary>
    public decimal LeftMargin
    {
        get { return this.GetXmlNodeDecimal(_leftMarginPath); }
        set
        {
            this.CreateMargins();
            this.SetXmlNodeString(_leftMarginPath, value.ToString(CultureInfo.InvariantCulture));
        }
    }

    const string _rightMarginPath = "d:pageMargins/@right";

    /// <summary>
    /// Right margin in inches
    /// </summary>
    public decimal RightMargin
    {
        get { return this.GetXmlNodeDecimal(_rightMarginPath); }
        set
        {
            this.CreateMargins();
            this.SetXmlNodeString(_rightMarginPath, value.ToString(CultureInfo.InvariantCulture));
        }
    }

    const string _topMarginPath = "d:pageMargins/@top";

    /// <summary>
    /// Top margin in inches
    /// </summary>
    public decimal TopMargin
    {
        get { return this.GetXmlNodeDecimal(_topMarginPath); }
        set
        {
            this.CreateMargins();
            this.SetXmlNodeString(_topMarginPath, value.ToString(CultureInfo.InvariantCulture));
        }
    }

    const string _bottomMarginPath = "d:pageMargins/@bottom";

    /// <summary>
    /// Bottom margin in inches
    /// </summary>
    public decimal BottomMargin
    {
        get { return this.GetXmlNodeDecimal(_bottomMarginPath); }
        set
        {
            this.CreateMargins();
            this.SetXmlNodeString(_bottomMarginPath, value.ToString(CultureInfo.InvariantCulture));
        }
    }

    const string _headerMarginPath = "d:pageMargins/@header";

    /// <summary>
    /// Header margin in inches
    /// </summary>
    public decimal HeaderMargin
    {
        get { return this.GetXmlNodeDecimal(_headerMarginPath); }
        set
        {
            this.CreateMargins();
            this.SetXmlNodeString(_headerMarginPath, value.ToString(CultureInfo.InvariantCulture));
        }
    }

    const string _footerMarginPath = "d:pageMargins/@footer";

    /// <summary>
    /// Footer margin in inches
    /// </summary>
    public decimal FooterMargin
    {
        get { return this.GetXmlNodeDecimal(_footerMarginPath); }
        set
        {
            this.CreateMargins();
            this.SetXmlNodeString(_footerMarginPath, value.ToString(CultureInfo.InvariantCulture));
        }
    }

    const string _orientationPath = "d:pageSetup/@orientation";

    /// <summary>
    /// Orientation 
    /// Portrait or Landscape
    /// </summary>
    public eOrientation Orientation
    {
        get { return (eOrientation)Enum.Parse(typeof(eOrientation), this.GetXmlNodeString(_orientationPath), true); }
        set { this.SetXmlNodeString(_orientationPath, value.ToString().ToLower(CultureInfo.InvariantCulture)); }
    }

    const string _fitToWidthPath = "d:pageSetup/@fitToWidth";

    /// <summary>
    /// Fit to Width in pages. 
    /// Set FitToPage to true when using this one. 
    /// 0 is automatic
    /// </summary>
    public int FitToWidth
    {
        get { return this.GetXmlNodeInt(_fitToWidthPath); }
        set { this.SetXmlNodeString(_fitToWidthPath, value.ToString()); }
    }

    const string _fitToHeightPath = "d:pageSetup/@fitToHeight";

    /// <summary>
    /// Fit to height in pages. 
    /// Set FitToPage to true when using this one. 
    /// 0 is automatic
    /// </summary>
    public int FitToHeight
    {
        get { return this.GetXmlNodeInt(_fitToHeightPath); }
        set { this.SetXmlNodeString(_fitToHeightPath, value.ToString()); }
    }

    const string _scalePath = "d:pageSetup/@scale";

    /// <summary>
    /// Print scale
    /// </summary>
    public int Scale
    {
        get { return this.GetXmlNodeInt(_scalePath); }
        set { this.SetXmlNodeString(_scalePath, value.ToString()); }
    }

    const string _fitToPagePath = "d:sheetPr/d:pageSetUpPr/@fitToPage";

    /// <summary>
    /// Fit To Page.
    /// </summary>
    public bool FitToPage
    {
        get { return this.GetXmlNodeBool(_fitToPagePath); }
        set { this.SetXmlNodeString(_fitToPagePath, value ? "1" : "0"); }
    }

    const string _headersPath = "d:printOptions/@headings";

    /// <summary>
    /// Print headings (column letter and row numbers)
    /// </summary>
    public bool ShowHeaders
    {
        get { return this.GetXmlNodeBool(_headersPath, false); }
        set { this.SetXmlNodeBool(_headersPath, value, false); }
    }

    /// <summary>
    /// Print titles
    /// Rows to be repeated after each pagebreak.
    /// The address must be a full row address (ex. 1:1)
    /// </summary>
    public ExcelAddress RepeatRows
    {
        get
        {
            if (this._ws.Names.ContainsKey("_xlnm.Print_Titles"))
            {
                ExcelRangeBase r = this._ws.Names["_xlnm.Print_Titles"] as ExcelRangeBase;

                if (r.Start.Column == 1 && r.End.Column == ExcelPackage.MaxColumns)
                {
                    return new ExcelAddress(r.FirstAddress);
                }
                else if (r._addresses != null)
                {
                    return new ExcelAddress(r._addresses[1].Address);
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }
        set
        {
            //Must span entire columns
            if (value.Addresses != null && !(value.Start.Column == 1 && value.End.Column == ExcelPackage.MaxColumns))
            {
                throw new
                    InvalidOperationException("Address must span full columns only (for ex. Address=\"A:A\" for the first column). Multi addresses is not allowed.");
            }

            ExcelAddress? vertAddr = this.RepeatColumns;
            string addr;

            if (vertAddr == null)
            {
                addr = value.Address;
            }
            else
            {
                addr = vertAddr.Address + "," + value.Address;
            }

            if (this._ws.Names.ContainsKey("_xlnm.Print_Titles"))
            {
                this._ws.Names["_xlnm.Print_Titles"].Address = addr;
            }
            else
            {
                _ = this._ws.Names.AddName("_xlnm.Print_Titles", new ExcelRangeBase(this._ws, addr));
            }
        }
    }

    /// <summary>
    /// Print titles
    /// Columns to be repeated after each pagebreak.
    /// The address must be a full column address (ex. A:A)
    /// </summary>
    public ExcelAddress RepeatColumns
    {
        get
        {
            if (this._ws.Names.ContainsKey("_xlnm.Print_Titles"))
            {
                ExcelRangeBase r = this._ws.Names["_xlnm.Print_Titles"] as ExcelRangeBase;

                if (r.Start.Row == 1 && r.End.Row == ExcelPackage.MaxRows)
                {
                    return new ExcelAddress(r.FirstAddress);
                }
                else if (r._addresses != null)
                {
                    return new ExcelAddress(r._addresses[1].Address);
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }
        set
        {
            //Must span entire rows
            if (value.Addresses != null && !(value.Start.Row == 1 && value.End.Row == ExcelPackage.MaxRows))
            {
                throw new InvalidOperationException("Address must span rows only (for ex. Address=\"1:1\" for the first row). Multi addresses is not allowed.");
            }

            ExcelAddress? horAddr = this.RepeatRows;
            string addr;

            if (horAddr == null)
            {
                addr = value.Address;
            }
            else
            {
                addr = value.Address + "," + horAddr.Address;
            }

            if (this._ws.Names.ContainsKey("_xlnm.Print_Titles"))
            {
                this._ws.Names["_xlnm.Print_Titles"].Address = addr;
            }
            else
            {
                _ = this._ws.Names.AddName("_xlnm.Print_Titles", new ExcelRangeBase(this._ws, addr));
            }
        }
    }

    /// <summary>
    /// The printarea.
    /// Null if no print area is set.
    /// </summary>
    public ExcelRangeBase PrintArea
    {
        get
        {
            if (this._ws.Names.ContainsKey("_xlnm.Print_Area"))
            {
                return this._ws.Names["_xlnm.Print_Area"];
            }
            else
            {
                return null;
            }
        }
        set
        {
            if (value == null)
            {
                this._ws.Names.Remove("_xlnm.Print_Area");
            }
            else if (this._ws.Names.ContainsKey("_xlnm.Print_Area"))
            {
                this._ws.Names["_xlnm.Print_Area"].Address = value.Address;
            }
            else
            {
                _ = this._ws.Names.AddName("_xlnm.Print_Area", value);
            }
        }
    }

    const string _gridLinesPath = "d:printOptions/@gridLines";

    /// <summary>
    /// Print gridlines 
    /// </summary>
    public bool ShowGridLines
    {
        get { return this.GetXmlNodeBool(_gridLinesPath, false); }
        set { this.SetXmlNodeBool(_gridLinesPath, value, false); }
    }

    const string _horizontalCenteredPath = "d:printOptions/@horizontalCentered";

    /// <summary>
    /// Horizontal centered when printing 
    /// </summary>w
    public bool HorizontalCentered
    {
        get { return this.GetXmlNodeBool(_horizontalCenteredPath, false); }
        set { this.SetXmlNodeBool(_horizontalCenteredPath, value, false); }
    }

    const string _verticalCenteredPath = "d:printOptions/@verticalCentered";

    /// <summary>
    /// Vertical centered when printing 
    /// </summary>
    public bool VerticalCentered
    {
        get { return this.GetXmlNodeBool(_verticalCenteredPath, false); }
        set { this.SetXmlNodeBool(_verticalCenteredPath, value, false); }
    }

    const string _pageOrderPath = "d:pageSetup/@pageOrder";

    /// <summary>
    /// Specifies printed page order
    /// </summary>
    public ePageOrder PageOrder
    {
        get
        {
            if (this.GetXmlNodeString(_pageOrderPath) == "overThenDown")
            {
                return ePageOrder.OverThenDown;
            }
            else
            {
                return ePageOrder.DownThenOver;
            }
        }
        set
        {
            if (value == ePageOrder.OverThenDown)
            {
                this.SetXmlNodeString(_pageOrderPath, "overThenDown");
            }
            else
            {
                this.DeleteNode(_pageOrderPath);
            }
        }
    }

    const string _blackAndWhitePath = "d:pageSetup/@blackAndWhite";

    /// <summary>
    /// Print black and white
    /// </summary>
    public bool BlackAndWhite
    {
        get { return this.GetXmlNodeBool(_blackAndWhitePath, false); }
        set { this.SetXmlNodeBool(_blackAndWhitePath, value, false); }
    }

    const string _draftPath = "d:pageSetup/@draft";

    /// <summary>
    /// Print a draft
    /// </summary>
    public bool Draft
    {
        get { return this.GetXmlNodeBool(_draftPath, false); }
        set { this.SetXmlNodeBool(_draftPath, value, false); }
    }

    const string _paperSizePath = "d:pageSetup/@paperSize";

    /// <summary>
    /// Paper size 
    /// </summary>
    public ePaperSize PaperSize
    {
        get
        {
            string s = this.GetXmlNodeString(_paperSizePath);

            if (s != "")
            {
                return (ePaperSize)int.Parse(s);
            }
            else
            {
                return ePaperSize.Letter;
            }
        }
        set { this.SetXmlNodeString(_paperSizePath, ((int)value).ToString()); }
    }

    /// <summary>
    /// All or none of the margin attributes must exist. Create all att ones.
    /// </summary>
    private void CreateMargins()
    {
        if (this._marginsCreated == false && this.TopNode.SelectSingleNode(_leftMarginPath, this.NameSpaceManager) == null)
        {
            this._marginsCreated = true;
            this.LeftMargin = 0.7087M;
            this.RightMargin = 0.7087M;
            this.TopMargin = 0.7480M;
            this.BottomMargin = 0.7480M;
            this.HeaderMargin = 0.315M;
            this.FooterMargin = 0.315M;
        }
    }
}