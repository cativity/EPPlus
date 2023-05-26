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

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Cell border style
    /// </summary>
    public sealed class ExcelBorderItem : StyleBase
    {
        eStyleClass _cls;
        StyleBase _parent;
        internal ExcelBorderItem (ExcelStyles styles, XmlHelper.ChangedEventHandler ChangedEvent, int worksheetID, string address, eStyleClass cls, StyleBase parent) : 
            base(styles, ChangedEvent, worksheetID, address)
	    {
            this._cls=cls;
            this._parent = parent;
	    }
        /// <summary>
        /// The line style of the border
        /// </summary>
        public ExcelBorderStyle Style
        {
            get
            {
                return this.GetSource().Style;
            }
            set
            {
                this._ChangedEvent(this, new StyleChangeEventArgs(this._cls, eStyleProperty.Style, value, this._positionID, this._address));
            }
        }
        ExcelColor _color=null;
        /// <summary>
        /// The color of the border
        /// </summary>
        public ExcelColor Color
        {
            get
            {
                if (this._color == null)
                {
                    this._color = new ExcelColor(this._styles, this._ChangedEvent, this._positionID, this._address, this._cls, this._parent);
                }
                return this._color;
            }
        }

        internal override string Id
        {
            get { return this.Style + this.Color.Id; }
        }
        internal override void SetIndex(int index)
        {
            this._parent.Index = index;
        }
        private ExcelBorderItemXml GetSource()
        {
            int ix = this._parent.Index < 0 ? 0 : this._parent.Index;

            switch(this._cls)
            {
                case eStyleClass.BorderTop:
                    return this._styles.Borders[ix].Top;
                case eStyleClass.BorderBottom:
                    return this._styles.Borders[ix].Bottom;
                case eStyleClass.BorderLeft:
                    return this._styles.Borders[ix].Left;
                case eStyleClass.BorderRight:
                    return this._styles.Borders[ix].Right;
                case eStyleClass.BorderDiagonal:
                    return this._styles.Borders[ix].Diagonal;
                default:
                    throw new Exception("Invalid class for Borderitem");
            }

        }
    }
}
