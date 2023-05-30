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
using System.Xml;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Style;

namespace OfficeOpenXml;

/// <summary>
/// Represents one or more columns within the worksheet
/// </summary>
public class ExcelColumn : IRangeID
{
    private ExcelWorksheet _worksheet;

    #region ExcelColumn Constructor

    /// <summary>
    /// Creates a new instance of the ExcelColumn class.  
    /// For internal use only!
    /// </summary>
    /// <param name="Worksheet"></param>
    /// <param name="col"></param>
    protected internal ExcelColumn(ExcelWorksheet Worksheet, int col)
    {
        this._worksheet = Worksheet;
        this._columnMin = col;
        this._columnMax = col;
        this._width = this._worksheet.DefaultColWidth;
    }

    #endregion

    internal int _columnMin;

    /// <summary>
    /// Sets the first column the definition refers to.
    /// </summary>
    public int ColumnMin
    {
        get { return this._columnMin; }

        //set { _columnMin=value; } 
    }

    internal int _columnMax;

    /// <summary>
    /// Sets the last column the definition refers to.
    /// </summary>
    public int ColumnMax
    {
        get { return this._columnMax; }
        set
        {
            if (value < this._columnMin && value > ExcelPackage.MaxColumns)
            {
                throw new Exception("ColumnMax out of range");
            }

            CellStoreEnumerator<ExcelValue>? cse = new CellStoreEnumerator<ExcelValue>(this._worksheet._values, 0, 0, 0, ExcelPackage.MaxColumns);

            while (cse.Next())
            {
                ExcelColumn? c = cse.Value._value as ExcelColumn;

                if (cse.Column > this._columnMin && c.ColumnMax <= value && cse.Column != this._columnMin)
                {
                    throw new Exception(string.Format("ColumnMax cannot span over existing column {0}.", c.ColumnMin));
                }
            }

            this._columnMax = value;
        }
    }

    /// <summary>
    /// Internal range id for the column
    /// </summary>
    internal ulong ColumnID
    {
        get { return GetColumnID(this._worksheet.SheetId, this.ColumnMin); }
    }

    #region ExcelColumn Hidden

    /// <summary>
    /// Allows the column to be hidden in the worksheet
    /// </summary>
    internal bool _hidden;

    /// <summary>
    /// Defines if the column is visible or hidden
    /// </summary>
    public bool Hidden
    {
        get { return this._hidden; }
        set
        {
            if (this._worksheet._package.DoAdjustDrawings)
            {
                double[,]? pos = this._worksheet.Drawings.GetDrawingWidths();
                this._hidden = value;
                this._worksheet.Drawings.AdjustWidth(pos);
            }
            else
            {
                this._hidden = value;
            }
        }
    }

    #endregion

    #region ExcelColumn Width

    internal double VisualWidth
    {
        get
        {
            if (this._hidden || (this.Collapsed && this.OutlineLevel > 0))
            {
                return 0;
            }
            else
            {
                return this._width;
            }
        }
    }

    internal double _width;

    /// <summary>
    /// Sets the width of the column in the worksheet
    /// </summary>
    public double Width
    {
        get { return this._width; }
        set
        {
            if (this._worksheet._package.DoAdjustDrawings)
            {
                double[,]? pos = this._worksheet.Drawings.GetDrawingWidths();
                this._width = value;
                this._worksheet.Drawings.AdjustWidth(pos);
            }
            else
            {
                this._width = value;
            }

            if (this._hidden && value != 0)
            {
                this._hidden = false;
            }
        }
    }

    /// <summary>
    /// If set to true a column automaticlly resize(grow wider) when a user inputs numbers in a cell. 
    /// </summary>
    public bool BestFit { get; set; }

    /// <summary>
    /// If the column is collapsed in outline mode
    /// </summary>
    public bool Collapsed { get; set; }

    /// <summary>
    /// Outline level. Zero if no outline
    /// </summary>
    public int OutlineLevel { get; set; }

    /// <summary>
    /// Phonetic
    /// </summary>
    public bool Phonetic { get; set; }

    #endregion

    #region ExcelColumn Style

    /// <summary>
    /// The Style applied to the whole column. Only effects cells with no individual style set. 
    /// Use Range object if you want to set specific styles.
    /// </summary>
    public ExcelStyle Style
    {
        get
        {
            string letter = ExcelCellBase.GetColumnLetter(this.ColumnMin);
            string endLetter = ExcelCellBase.GetColumnLetter(this.ColumnMax);

            return this._worksheet.Workbook.Styles.GetStyleObject(this.StyleID, this._worksheet.PositionId, letter + ":" + endLetter);
        }
    }

    internal string _styleName = "";

    /// <summary>
    /// Sets the style for the entire column using a style name.
    /// </summary>
    public string StyleName
    {
        get { return this._styleName; }
        set
        {
            this.StyleID = this._worksheet.Workbook.Styles.GetStyleIdFromName(value);
            this._styleName = value;
        }
    }

    /// <summary>
    /// Sets the style for the entire column using the style ID.           
    /// </summary>
    public int StyleID
    {
        get { return this._worksheet.GetStyleInner(0, this.ColumnMin); }
        set { this._worksheet.SetStyleInner(0, this.ColumnMin, value); }
    }

    /// <summary>
    /// Adds a manual page break after the column.
    /// </summary>
    public bool PageBreak { get; set; }

    /// <summary>
    /// Merges all cells of the column
    /// </summary>
    public bool Merged
    {
        get { return this._worksheet.MergedCells[0, this.ColumnMin] != null; }
        set { this._worksheet.MergedCells.Add(new ExcelAddressBase(1, this.ColumnMin, ExcelPackage.MaxRows, this.ColumnMax), true); }
    }

    #endregion

    /// <summary>
    /// Returns the range of columns covered by the column definition.
    /// </summary>
    /// <returns>A string describing the range of columns covered by the column definition.</returns>
    public override string ToString()
    {
        return string.Format("Column Range: {0} to {1}", this.ColumnMin, this.ColumnMax);
    }

    /// <summary>
    /// Set the column width from the content of the range. The minimum width is the value of the ExcelWorksheet.defaultColumnWidth property.
    /// Note: Cells containing formulas are ignored unless a calculation is performed.
    ///       Wrapped and merged cells are also ignored.
    /// </summary>
    public void AutoFit()
    {
        this._worksheet.Cells[1, this._columnMin, ExcelPackage.MaxRows, this._columnMax].AutoFitColumns();
    }

    /// <summary>
    /// Set the column width from the content.
    /// Note: Cells containing formulas are ignored unless a calculation is performed.
    ///       Wrapped and merged cells are also ignored.
    /// </summary>
    /// <param name="MinimumWidth">Minimum column width</param>
    public void AutoFit(double MinimumWidth)
    {
        this._worksheet.Cells[1, this._columnMin, ExcelPackage.MaxRows, this._columnMax].AutoFitColumns(MinimumWidth);
    }

    /// <summary>
    /// Set the column width from the content.
    /// Note: Cells containing formulas are ignored unless a calculation is performed.
    ///       Wrapped and merged cells are also ignored.
    /// </summary>
    /// <param name="MinimumWidth">Minimum column width</param>
    /// <param name="MaximumWidth">Maximum column width</param>
    public void AutoFit(double MinimumWidth, double MaximumWidth)
    {
        this._worksheet.Cells[1, this._columnMin, ExcelPackage.MaxRows, this._columnMax].AutoFitColumns(MinimumWidth, MaximumWidth);
    }

    /// <summary>
    /// Get the internal RangeID
    /// </summary>
    /// <param name="sheetID">Sheet no</param>
    /// <param name="column">Column</param>
    /// <returns></returns>
    internal static ulong GetColumnID(int sheetID, int column)
    {
        return (ulong)sheetID + ((ulong)column << 15);
    }

    internal static int ColumnWidthToPixels(decimal columnWidth, decimal mdw)
    {
        return (int)decimal.Truncate(((256 * columnWidth) + decimal.Truncate(128 / mdw)) / 256 * mdw);
    }

    #region IRangeID Members

    ulong IRangeID.RangeID
    {
        get { return this.ColumnID; }
        set
        {
            int prevColMin = this._columnMin;
            this._columnMin = (int)(value >> 15) & 0x3FF;
            this._columnMax += prevColMin - this.ColumnMin;

            //Todo:More Validation
            if (this._columnMax > ExcelPackage.MaxColumns)
            {
                this._columnMax = ExcelPackage.MaxColumns;
            }
        }
    }

    #endregion

    /// <summary>
    /// Copies the current column to a new worksheet
    /// </summary>
    /// <param name="added">The worksheet where the copy will be created</param>
    internal ExcelColumn Clone(ExcelWorksheet added)
    {
        return this.Clone(added, this.ColumnMin);
    }

    internal ExcelColumn Clone(ExcelWorksheet added, int col)
    {
        ExcelColumn newCol = added.Column(col);
        newCol.ColumnMax = this.ColumnMax;
        newCol.BestFit = this.BestFit;
        newCol.Collapsed = this.Collapsed;
        newCol.OutlineLevel = this.OutlineLevel;
        newCol.PageBreak = this.PageBreak;
        newCol.Phonetic = this.Phonetic;
        newCol._styleName = this._styleName;
        newCol.StyleID = this.StyleID;
        newCol.Width = this.Width;
        newCol.Hidden = this.Hidden;

        return newCol;
    }
}