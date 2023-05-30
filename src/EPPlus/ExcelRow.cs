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
using OfficeOpenXml.Style;

namespace OfficeOpenXml;

internal class RowInternal
{
    internal double Height = -1;
    internal bool Hidden;
    internal bool Collapsed;
    internal short OutlineLevel;
    internal bool PageBreak;
    internal bool Phonetic;
    internal bool CustomHeight;
    internal int MergeID;

    internal RowInternal Clone() =>
        new()
        {
            Height = this.Height,
            Hidden = this.Hidden,
            Collapsed = this.Collapsed,
            OutlineLevel = this.OutlineLevel,
            PageBreak = this.PageBreak,
            Phonetic = this.Phonetic,
            CustomHeight = this.CustomHeight,
            MergeID = this.MergeID
        };
}

/// <summary>
/// Represents an individual row in the spreadsheet.
/// </summary>
public class ExcelRow : IRangeID
{
    private ExcelWorksheet _worksheet;
    private XmlElement _rowElement = null;

    /// <summary>
    /// Internal RowID.
    /// </summary>
    [Obsolete]
    public ulong RowID => GetRowID(this._worksheet.SheetId, this.Row);

    #region ExcelRow Constructor

    /// <summary>
    /// Creates a new instance of the ExcelRow class. 
    /// For internal use only!
    /// </summary>
    /// <param name="Worksheet">The parent worksheet</param>
    /// <param name="row">The row number</param>
    internal ExcelRow(ExcelWorksheet Worksheet, int row)
    {
        this._worksheet = Worksheet;
        this.Row = row;
    }

    #endregion

    /// <summary>
    /// Provides access to the node representing the row.
    /// </summary>
    internal XmlNode Node => this._rowElement;

    #region ExcelRow Hidden

    /// <summary>
    /// Allows the row to be hidden in the worksheet
    /// </summary>
    public bool Hidden
    {
        get
        {
            RowInternal? r = (RowInternal)this._worksheet.GetValueInner(this.Row, 0);

            if (r == null)
            {
                return false;
            }
            else
            {
                return r.Hidden;
            }
        }
        set
        {
            RowInternal? r = this.GetRowInternal();
            r.Hidden = value;
        }
    }

    #endregion

    #region ExcelRow Height

    /// <summary>
    /// Sets the height of the row
    /// </summary>
    public double Height
    {
        get
        {
            RowInternal? r = (RowInternal)this._worksheet.GetValueInner(this.Row, 0);

            if (r == null || r.Height < 0)
            {
                return this._worksheet.DefaultRowHeight;
            }
            else
            {
                return r.Height;
            }
        }
        set
        {
            RowInternal? r = this.GetRowInternal();

            if (this._worksheet._package.DoAdjustDrawings)
            {
                double[,]? pos = this._worksheet.Drawings.GetDrawingHeight(); //Fixes issue 14846
                _ = this._worksheet.RowHeightCache.Remove(this.Row - 1);
                r.Height = value;
                this._worksheet.Drawings.AdjustHeight(pos);
            }
            else
            {
                r.Height = value;
            }

            if (r.Hidden && value != 0)
            {
                this.Hidden = false;
            }

            r.CustomHeight = true;
        }
    }

    /// <summary>
    /// Set to true if You don't want the row to Autosize
    /// </summary>
    public bool CustomHeight
    {
        get
        {
            RowInternal? r = (RowInternal)this._worksheet.GetValueInner(this.Row, 0);

            if (r == null)
            {
                return false;
            }
            else
            {
                return r.CustomHeight;
            }
        }
        set
        {
            RowInternal? r = this.GetRowInternal();
            r.CustomHeight = value;
        }
    }

    #endregion

    internal string _styleName = "";

    /// <summary>
    /// Sets the style for the entire column using a style name.
    /// </summary>
    public string StyleName
    {
        get => this._styleName;
        set
        {
            this.StyleID = this._worksheet.Workbook.Styles.GetStyleIdFromName(value);
            this._styleName = value;
        }
    }

    /// <summary>
    /// Sets the style for the entire row using the style ID.  
    /// </summary>
    public int StyleID
    {
        get => this._worksheet.GetStyleInner(this.Row, 0);
        set => this._worksheet.SetStyleInner(this.Row, 0, value);
    }

    /// <summary>
    /// Rownumber
    /// </summary>
    public int Row { get; set; }

    /// <summary>
    /// If outline level is set this tells that the row is collapsed
    /// </summary>
    public bool Collapsed
    {
        get
        {
            RowInternal? r = (RowInternal)this._worksheet.GetValueInner(this.Row, 0);

            if (r == null)
            {
                return false;
            }
            else
            {
                return r.Collapsed;
            }
        }
        set
        {
            RowInternal? r = this.GetRowInternal();
            r.Collapsed = value;
        }
    }

    /// <summary>
    /// Outline level.
    /// </summary>
    public int OutlineLevel
    {
        get
        {
            RowInternal? r = (RowInternal)this._worksheet.GetValueInner(this.Row, 0);

            if (r == null)
            {
                return 0;
            }
            else
            {
                return r.OutlineLevel;
            }
        }
        set
        {
            RowInternal? r = this.GetRowInternal();
            r.OutlineLevel = (short)value;
        }
    }

    private RowInternal GetRowInternal() => GetRowInternal(this._worksheet, this.Row);

    internal static RowInternal GetRowInternal(ExcelWorksheet ws, int row)
    {
        RowInternal? r = (RowInternal)ws.GetValueInner(row, 0);

        if (r == null)
        {
            r = new RowInternal();
            ws.SetValueInner(row, 0, r);
        }

        return r;
    }

    /// <summary>
    /// Show phonetic Information
    /// </summary>
    public bool Phonetic
    {
        get
        {
            RowInternal? r = (RowInternal)this._worksheet.GetValueInner(this.Row, 0);

            if (r == null)
            {
                return false;
            }
            else
            {
                return r.Phonetic;
            }
        }
        set
        {
            RowInternal? r = this.GetRowInternal();
            r.Phonetic = value;
        }
    }

    /// <summary>
    /// The Style applied to the whole row. Only effekt cells with no individual style set. 
    /// Use the <see cref="ExcelWorksheet.Cells"/> Style property if you want to set specific styles.
    /// </summary>
    public ExcelStyle Style =>
        this._worksheet.Workbook.Styles.GetStyleObject(this.StyleID,
                                                       this._worksheet.PositionId,
                                                       this.Row.ToString(CultureInfo.InvariantCulture)
                                                       + ":"
                                                       + this.Row.ToString(CultureInfo.InvariantCulture));

    /// <summary>
    /// Adds a manual page break after the row.
    /// </summary>
    public bool PageBreak
    {
        get
        {
            RowInternal? r = (RowInternal)this._worksheet.GetValueInner(this.Row, 0);

            if (r == null)
            {
                return false;
            }
            else
            {
                return r.PageBreak;
            }
        }
        set
        {
            RowInternal? r = this.GetRowInternal();
            r.PageBreak = value;
        }
    }

    /// <summary>
    /// Merge all cells in the row
    /// </summary>
    public bool Merged
    {
        get => this._worksheet.MergedCells[this.Row, 0] != null;
        set => this._worksheet.MergedCells.Add(new ExcelAddressBase(this.Row, 1, this.Row, ExcelPackage.MaxColumns), true);
    }

    internal static ulong GetRowID(int sheetID, int row) => (ulong)sheetID + ((ulong)row << 29);

    #region IRangeID Members

    [Obsolete]
    ulong IRangeID.RangeID
    {
        get => this.RowID;
        set => this.Row = (int)(value >> 29);
    }

    #endregion

    /// <summary>
    /// Copies the current row to a new worksheet
    /// </summary>
    /// <param name="added">The worksheet where the copy will be created</param>
    internal void Clone(ExcelWorksheet added)
    {
        RowInternal? rowSource = this._worksheet.GetValue(this.Row, 0) as RowInternal;

        if (rowSource != null)
        {
            added.SetValueInner(this.Row, 0, rowSource.Clone());
        }
    }
}