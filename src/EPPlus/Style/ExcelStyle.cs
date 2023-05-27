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
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Style;

/// <summary>
/// Toplevel class for cell styling
/// </summary>
public sealed class ExcelStyle : StyleBase
{
    ExcelXfs _xfs;

    internal ExcelStyle(ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int positionID, string Address, int xfsId)
        : base(styles, ChangedEvent, positionID, Address)
    {
        this.Index = xfsId;
        this.Styles = styles;
        this.PositionID = positionID;

        if (positionID > -1)
        {
            if (xfsId == 0)
            {
                int id = this._styles.NamedStyles.FindIndexByBuildInId(0);

                if (id > -1 && id < this._styles.CellStyleXfs.Count)
                {
                    this._xfs = this._styles.CellStyleXfs[this._styles.NamedStyles[id].StyleXfId];
                }
                else
                {
                    this._xfs = this._styles.CellXfs[0];
                }
            }
            else
            {
                this._xfs = this._styles.CellXfs[xfsId];
            }
        }
        else
        {
            if (this._styles.CellStyleXfs.Count == 0) //CellStyleXfs.Count should never be 0, but for some custom build sheets this can happend.
            {
                ExcelXfs? item = this._styles.CellXfs[0].Copy();
                _ = this._styles.CellStyleXfs.Add(item.Id, item);
            }

            this._xfs = this._styles.CellStyleXfs[xfsId];
        }

        this.Numberformat = new ExcelNumberFormat(styles, ChangedEvent, this.PositionID, Address, this._xfs.NumberFormatId);
        this.Font = new ExcelFont(styles, ChangedEvent, this.PositionID, Address, this._xfs.FontId);
        this.Fill = new ExcelFill(styles, ChangedEvent, this.PositionID, Address, this._xfs.FillId);
        this.Border = new Border(styles, ChangedEvent, this.PositionID, Address, this._xfs.BorderId);
    }

    /// <summary>
    /// Numberformat
    /// </summary>
    public ExcelNumberFormat Numberformat { get; set; }

    /// <summary>
    /// Font styling
    /// </summary>
    public ExcelFont Font { get; set; }

    /// <summary>
    /// Fill Styling
    /// </summary>
    public ExcelFill Fill { get; set; }

    /// <summary>
    /// Border 
    /// </summary>
    public Border Border { get; set; }

    /// <summary>
    /// The horizontal alignment in the cell
    /// </summary>
    public ExcelHorizontalAlignment HorizontalAlignment
    {
        get { return this._xfs.HorizontalAlignment; }
        set { _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.HorizontalAlign, value, this._positionID, this._address)); }
    }

    /// <summary>
    /// The vertical alignment in the cell
    /// </summary>
    public ExcelVerticalAlignment VerticalAlignment
    {
        get
        {
            //return _styles.CellXfs[Index].VerticalAlignment;
            return this._xfs.VerticalAlignment;
        }
        set { _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.VerticalAlign, value, this._positionID, this._address)); }
    }

    /// <summary>
    /// If the cells justified or distributed alignment should be used on the last line of text.
    /// </summary>
    public bool JustifyLastLine
    {
        get { return this._xfs.JustifyLastLine; }
        set { _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.JustifyLastLine, value, this._positionID, this._address)); }
    }

    /// <summary>
    /// Wrap the text
    /// </summary>
    public bool WrapText
    {
        get { return this._xfs.WrapText; }
        set { _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.WrapText, value, this._positionID, this._address)); }
    }

    /// <summary>
    /// Readingorder
    /// </summary>
    public ExcelReadingOrder ReadingOrder
    {
        get { return this._xfs.ReadingOrder; }
        set { _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.ReadingOrder, value, this._positionID, this._address)); }
    }

    /// <summary>
    /// Makes the text vertically. This is the same as setting <see cref="TextRotation"/> to 255.
    /// </summary>
    public void SetTextVertical()
    {
        this.TextRotation = 255;
    }

    /// <summary>
    /// Shrink the text to fit
    /// </summary>
    public bool ShrinkToFit
    {
        get { return this._xfs.ShrinkToFit; }
        set { _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.ShrinkToFit, value, this._positionID, this._address)); }
    }

    /// <summary>
    /// The margin between the border and the text
    /// </summary>
    public int Indent
    {
        get { return this._xfs.Indent; }
        set
        {
            if (value < 0 || value > 250)
            {
                throw new ArgumentOutOfRangeException("Indent must be between 0 and 250");
            }

            _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.Indent, value, this._positionID, this._address));
        }
    }

    /// <summary>
    /// Text orientation in degrees. Values range from 0 to 180 or 255. 
    /// Setting the rotation to 255 will align text vertically.
    /// </summary>
    public int TextRotation
    {
        get { return this._xfs.TextRotation; }
        set
        {
            if ((value < 0 || value > 180) && value != 255)
            {
                throw new ArgumentOutOfRangeException("TextRotation out of range.");
            }

            _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.TextRotation, value, this._positionID, this._address));
        }
    }

    /// <summary>
    /// If true the cell is locked for editing when the sheet is protected
    /// <seealso cref="ExcelWorksheet.Protection"/>
    /// </summary>
    public bool Locked
    {
        get { return this._xfs.Locked; }
        set { _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.Locked, value, this._positionID, this._address)); }
    }

    /// <summary>
    /// If true the formula is hidden when the sheet is protected.
    /// <seealso cref="ExcelWorksheet.Protection"/>
    /// </summary>
    public bool Hidden
    {
        get { return this._xfs.Hidden; }
        set { _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.Hidden, value, this._positionID, this._address)); }
    }

    /// <summary>
    /// If true the cell has a quote prefix, which indicates the value of the cell is text.
    /// </summary>
    public bool QuotePrefix
    {
        get { return this._xfs.QuotePrefix; }
        set { _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.QuotePrefix, value, this._positionID, this._address)); }
    }

    const string xfIdPath = "@xfId";

    /// <summary>
    /// The index in the style collection
    /// </summary>
    public int XfId
    {
        get { return this._xfs.XfId; }
        set { _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Style, eStyleProperty.XfId, value, this._positionID, this._address)); }
    }

    internal int PositionID { get; set; }

    internal ExcelStyles Styles { get; set; }

    internal override string Id
    {
        get
        {
            return this.Numberformat.Id
                   + "|"
                   + this.Font.Id
                   + "|"
                   + this.Fill.Id
                   + "|"
                   + this.Border.Id
                   + "|"
                   + this.VerticalAlignment
                   + "|"
                   + this.HorizontalAlignment
                   + "|"
                   + this.WrapText.ToString()
                   + "|"
                   + this.ReadingOrder.ToString()
                   + "|"
                   + this.XfId.ToString()
                   + "|"
                   + this.QuotePrefix.ToString();
        }
    }
}