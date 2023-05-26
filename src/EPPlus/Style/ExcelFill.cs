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
using System.Drawing;
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// The background fill of a cell
    /// </summary>
    public class ExcelFill : StyleBase
    {
        internal ExcelFill(ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index) :
            base(styles, ChangedEvent, PositionID, address)

        {
            this.Index = index;
        }
        /// <summary>
        /// The pattern for solid fills.
        /// </summary>
        public ExcelFillStyle PatternType
        {
            get
            {
                if (this.Index == int.MinValue)
                {
                    return ExcelFillStyle.None;
                }
                else
                {
                    return this._styles.Fills[this.Index].PatternType;
                }
            }
            set
            {
                if (this._gradient != null)
                {
                    this._gradient = null;
                }

                this._ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Fill, eStyleProperty.PatternType, value, this._positionID, this._address));
            }
        }
        ExcelColor _patternColor = null;
        /// <summary>
        /// The color of the pattern
        /// </summary>
        public ExcelColor PatternColor
        {
            get
            {
                if (this._patternColor == null)
                {
                    this._patternColor = new ExcelColor(this._styles, this._ChangedEvent, this._positionID, this._address, eStyleClass.FillPatternColor, this);
                    if (this._gradient != null)
                    {
                        this._gradient = null;
                    }
                }
                return this._patternColor;
            }
        }
        ExcelColor _backgroundColor = null;
        /// <summary>
        /// The background color
        /// </summary>
        public ExcelColor BackgroundColor
        {
            get
            {
                if (this._backgroundColor == null)
                {
                    this._backgroundColor = new ExcelColor(this._styles, this._ChangedEvent, this._positionID, this._address, eStyleClass.FillBackgroundColor, this);
                    if (this._gradient != null)
                    {
                        this._gradient = null;
                    }
                }
                return this._backgroundColor;
                
            }
        }
        ExcelGradientFill _gradient=null;
        /// <summary>
        /// Access to properties for gradient fill.
        /// </summary>
        public ExcelGradientFill Gradient 
        {
            get
            {
                if (this._gradient == null)
                {
                    this._gradient = new ExcelGradientFill(this._styles, this._ChangedEvent, this._positionID, this._address, this.Index);
                    this._backgroundColor = null;
                    this._patternColor = null;
                }
                return this._gradient;
            }
        }
        internal override string Id
        {
            get
            {
                if (this._gradient == null)
                {
                    return this.PatternType + this.PatternColor.Id + this.BackgroundColor.Id;
                }
                else
                {
                    return this._gradient.Id;
                }
            }
        }
        /// <summary>
        /// Set the background to a specific color and fillstyle
        /// </summary>
        /// <param name="color">the color</param>
        /// <param name="fillStyle">The fillstyle. Default Solid</param>
        public void SetBackground(Color color, ExcelFillStyle fillStyle=ExcelFillStyle.Solid)
        {
            this.PatternType = fillStyle;
            this.BackgroundColor.SetColor(color);
        }
        /// <summary>
        /// Set the background to a specific color and fillstyle
        /// </summary>
        /// <param name="color">The indexed color</param>
        /// <param name="fillStyle">The fillstyle. Default Solid</param>
        public void SetBackground(ExcelIndexedColor color, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            this.PatternType = fillStyle;
            this.BackgroundColor.SetColor(color);
        }
        /// <summary>
        /// Set the background to a specific color and fillstyle
        /// </summary>
        /// <param name="color">The theme color</param>
        /// <param name="fillStyle">The fillstyle. Default Solid</param>
        public void SetBackground(eThemeSchemeColor color, ExcelFillStyle fillStyle = ExcelFillStyle.Solid)
        {
            this.PatternType = fillStyle;
            this.BackgroundColor.SetColor(color);
        }
    }
}
    