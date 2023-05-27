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
using System.Globalization;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Style;

/// <summary>
/// The background fill of a cell
/// </summary>
public class ExcelGradientFill : StyleBase
{
    internal ExcelGradientFill(ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index)
        : base(styles, ChangedEvent, PositionID, address)

    {
        this.Index = index;
    }

    /// <summary>
    /// Angle of the linear gradient
    /// </summary>
    public double Degree
    {
        get
        {
            if (this.Index < 0)
            {
                return 0;
            }

            return ((ExcelGradientFillXml)this._styles.Fills[this.Index]).Degree;
        }
        set
        {
            _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientDegree, value, this._positionID, this._address));
        }
    }

    /// <summary>
    /// Linear or Path gradient
    /// </summary>
    public ExcelFillGradientType Type
    {
        get
        {
            if (this.Index < 0)
            {
                return ExcelFillGradientType.None;
            }

            return ((ExcelGradientFillXml)this._styles.Fills[this.Index]).Type;
        }
        set
        {
            _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientType, value, this._positionID, this._address));
        }
    }

    /// <summary>
    /// The top position of the inner rectangle (color 1) in percentage format (from the top to the bottom). 
    /// Spans from 0 to 1
    /// </summary>
    public double Top
    {
        get
        {
            if (this.Index < 0)
            {
                return 0;
            }

            return ((ExcelGradientFillXml)this._styles.Fills[this.Index]).Top;
        }
        set
        {
            if ((value < 0) | (value > 1))
            {
                throw new ArgumentOutOfRangeException("Value must be between 0 and 1");
            }

            _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientTop, value, this._positionID, this._address));
        }
    }

    /// <summary>
    /// The bottom position of the inner rectangle (color 1) in percentage format (from the top to the bottom). 
    /// Spans from 0 to 1
    /// </summary>
    public double Bottom
    {
        get
        {
            if (this.Index < 0)
            {
                return 0;
            }

            return ((ExcelGradientFillXml)this._styles.Fills[this.Index]).Bottom;
        }
        set
        {
            if ((value < 0) | (value > 1))
            {
                throw new ArgumentOutOfRangeException("Value must be between 0 and 1");
            }

            _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientBottom, value, this._positionID, this._address));
        }
    }

    /// <summary>
    /// The left position of the inner rectangle (color 1) in percentage format (from the left to the right). 
    /// Spans from 0 to 1
    /// </summary>
    public double Left
    {
        get
        {
            if (this.Index < 0)
            {
                return 0;
            }

            return ((ExcelGradientFillXml)this._styles.Fills[this.Index]).Left;
        }
        set
        {
            if ((value < 0) | (value > 1))
            {
                throw new ArgumentOutOfRangeException("Value must be between 0 and 1");
            }

            _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientLeft, value, this._positionID, this._address));
        }
    }

    /// <summary>
    /// The right position of the inner rectangle (color 1) in percentage format (from the left to the right). 
    /// Spans from 0 to 1
    /// </summary>
    public double Right
    {
        get
        {
            if (this.Index < 0)
            {
                return 0;
            }

            return ((ExcelGradientFillXml)this._styles.Fills[this.Index]).Right;
        }
        set
        {
            if ((value < 0) | (value > 1))
            {
                throw new ArgumentOutOfRangeException("Value must be between 0 and 1");
            }

            _ = this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientRight, value, this._positionID, this._address));
        }
    }

    ExcelColor _gradientColor1 = null;

    /// <summary>
    /// Gradient Color 1
    /// </summary>
    public ExcelColor Color1
    {
        get
        {
            return this._gradientColor1 ??= new ExcelColor(this._styles,
                                                           this._ChangedEvent,
                                                           this._positionID,
                                                           this._address,
                                                           eStyleClass.FillGradientColor1,
                                                           this);
        }
    }

    ExcelColor _gradientColor2 = null;

    /// <summary>
    /// Gradient Color 2
    /// </summary>
    public ExcelColor Color2
    {
        get
        {
            return this._gradientColor2 ??= new ExcelColor(this._styles,
                                                           this._ChangedEvent,
                                                           this._positionID,
                                                           this._address,
                                                           eStyleClass.FillGradientColor2,
                                                           this);
        }
    }

    internal override string Id
    {
        get
        {
            return this.Degree.ToString()
                   + this.Type
                   + this.Color1.Id
                   + this.Color2.Id
                   + this.Top.ToString()
                   + this.Bottom.ToString()
                   + this.Left.ToString()
                   + this.Right.ToString();
        }
    }
}