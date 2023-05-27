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
using System.Drawing;
using System.Text;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Style;

/// <summary>
/// Cell Border style
/// </summary>
public sealed class Border : StyleBase
{
    internal Border(ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index)
        : base(styles, ChangedEvent, PositionID, address)
    {
        this.Index = index;
    }

    /// <summary>
    /// Left border style
    /// </summary>
    public ExcelBorderItem Left
    {
        get { return new ExcelBorderItem(this._styles, this._ChangedEvent, this._positionID, this._address, eStyleClass.BorderLeft, this); }
    }

    /// <summary>
    /// Right border style
    /// </summary>
    public ExcelBorderItem Right
    {
        get { return new ExcelBorderItem(this._styles, this._ChangedEvent, this._positionID, this._address, eStyleClass.BorderRight, this); }
    }

    /// <summary>
    /// Top border style
    /// </summary>
    public ExcelBorderItem Top
    {
        get { return new ExcelBorderItem(this._styles, this._ChangedEvent, this._positionID, this._address, eStyleClass.BorderTop, this); }
    }

    /// <summary>
    /// Bottom border style
    /// </summary>
    public ExcelBorderItem Bottom
    {
        get { return new ExcelBorderItem(this._styles, this._ChangedEvent, this._positionID, this._address, eStyleClass.BorderBottom, this); }
    }

    /// <summary>
    /// 0Diagonal border style
    /// </summary>
    public ExcelBorderItem Diagonal
    {
        get { return new ExcelBorderItem(this._styles, this._ChangedEvent, this._positionID, this._address, eStyleClass.BorderDiagonal, this); }
    }

    /// <summary>
    /// A diagonal from the bottom left to top right of the cell
    /// </summary>
    public bool DiagonalUp
    {
        get
        {
            if (this.Index >= 0)
            {
                return this._styles.Borders[this.Index].DiagonalUp;
            }
            else
            {
                return false;
            }
        }
        set { this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Border, eStyleProperty.BorderDiagonalUp, value, this._positionID, this._address)); }
    }

    /// <summary>
    /// A diagonal from the top left to bottom right of the cell
    /// </summary>
    public bool DiagonalDown
    {
        get
        {
            if (this.Index >= 0)
            {
                return this._styles.Borders[this.Index].DiagonalDown;
            }
            else
            {
                return false;
            }
        }
        set
        {
            this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Border, eStyleProperty.BorderDiagonalDown, value, this._positionID, this._address));
        }
    }

    internal override string Id
    {
        get { return this.Top.Id + this.Bottom.Id + this.Left.Id + this.Right.Id + this.Diagonal.Id + this.DiagonalUp + this.DiagonalDown; }
    }

    /// <summary>
    /// Set the border style around the range.
    /// </summary>
    /// <param name="Style">The border style</param>
    public void BorderAround(ExcelBorderStyle Style)
    {
        this.BorderAround(Style, Color.Empty);
    }

    /// <summary>
    /// Set the border style around the range.
    /// </summary>
    /// <param name="Style">The border style</param>
    /// <param name="Color">The color of the border</param>
    public void BorderAround(ExcelBorderStyle Style, Color Color)
    {
        ExcelAddressBase? addr = new ExcelAddressBase(this._address);

        if (addr.Addresses?.Count > 1)
        {
            foreach (ExcelAddressBase? a in addr.Addresses)
            {
                this.SetBorderAroundStyle(Style, a);

                if (!Color.IsEmpty)
                {
                    this.SetBorderColor(Color, a);
                }
            }
        }
        else
        {
            this.SetBorderAroundStyle(Style, addr);

            if (!Color.IsEmpty)
            {
                this.SetBorderColor(Color, addr);
            }
        }
    }

    private void SetBorderColor(Color Color, ExcelAddressBase addr)
    {
        this._ChangedEvent(this,
                           new StyleChangeEventArgs(eStyleClass.BorderTop,
                                                    eStyleProperty.Color,
                                                    Color.ToArgb().ToString("X"),
                                                    this._positionID,
                                                    new ExcelAddress(addr._fromRow, addr._fromCol, addr._fromRow, addr._toCol).Address));

        this._ChangedEvent(this,
                           new StyleChangeEventArgs(eStyleClass.BorderBottom,
                                                    eStyleProperty.Color,
                                                    Color.ToArgb().ToString("X"),
                                                    this._positionID,
                                                    new ExcelAddress(addr._toRow, addr._fromCol, addr._toRow, addr._toCol).Address));

        this._ChangedEvent(this,
                           new StyleChangeEventArgs(eStyleClass.BorderLeft,
                                                    eStyleProperty.Color,
                                                    Color.ToArgb().ToString("X"),
                                                    this._positionID,
                                                    new ExcelAddress(addr._fromRow, addr._fromCol, addr._toRow, addr._fromCol).Address));

        this._ChangedEvent(this,
                           new StyleChangeEventArgs(eStyleClass.BorderRight,
                                                    eStyleProperty.Color,
                                                    Color.ToArgb().ToString("X"),
                                                    this._positionID,
                                                    new ExcelAddress(addr._fromRow, addr._toCol, addr._toRow, addr._toCol).Address));
    }

    private void SetBorderAroundStyle(ExcelBorderStyle Style, ExcelAddressBase addr)
    {
        this._ChangedEvent(this,
                           new StyleChangeEventArgs(eStyleClass.BorderTop,
                                                    eStyleProperty.Style,
                                                    Style,
                                                    this._positionID,
                                                    new ExcelAddress(addr._fromRow, addr._fromCol, addr._fromRow, addr._toCol).Address));

        this._ChangedEvent(this,
                           new StyleChangeEventArgs(eStyleClass.BorderBottom,
                                                    eStyleProperty.Style,
                                                    Style,
                                                    this._positionID,
                                                    new ExcelAddress(addr._toRow, addr._fromCol, addr._toRow, addr._toCol).Address));

        this._ChangedEvent(this,
                           new StyleChangeEventArgs(eStyleClass.BorderLeft,
                                                    eStyleProperty.Style,
                                                    Style,
                                                    this._positionID,
                                                    new ExcelAddress(addr._fromRow, addr._fromCol, addr._toRow, addr._fromCol).Address));

        this._ChangedEvent(this,
                           new StyleChangeEventArgs(eStyleClass.BorderRight,
                                                    eStyleProperty.Style,
                                                    Style,
                                                    this._positionID,
                                                    new ExcelAddress(addr._fromRow, addr._toCol, addr._toRow, addr._toCol).Address));
    }
}