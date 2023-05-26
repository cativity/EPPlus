/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    10/21/2020         EPPlus Software AB           Controls 
 *************************************************************************************************/
using OfficeOpenXml.Packaging;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    /// <summary>
    /// Represents a scroll bar form control
    /// </summary>
    public class ExcelControlScrollBar : ExcelControl
    {
        internal ExcelControlScrollBar(ExcelDrawings drawings, XmlElement drawNode, string name, ExcelGroupShape parent=null) : base(drawings, drawNode, name, parent)
        {
            this.SetSize(30, 150); //Default size
        }
        internal ExcelControlScrollBar(ExcelDrawings drawings, XmlNode drawNode, ControlInternal control, ZipPackagePart part, XmlDocument controlPropertiesXml, ExcelGroupShape parent = null)
            : base(drawings, drawNode, control, part, controlPropertiesXml, parent)
        {
        }

        /// <summary>
        /// The type of form control
        /// </summary>
        public override eControlType ControlType => eControlType.ScrollBar;

        /// <summary>
        /// Gets or sets if scroll bar is horizontal or vertical
        /// </summary>
        public bool Horizontal
        {
            get
            {
                return this._ctrlProp.GetXmlNodeBool("@horiz");
            }
            set
            {
                this._ctrlProp.SetXmlNodeBool("@horiz", value);
                if(value)
                {
                    this._vmlProp.CreateNode("x:Horiz");
                }
                else
                {
                    this._vmlProp.DeleteNode("x:Horiz");
                }
            }
        }
        /// <summary>
        /// How much the scroll bar is incremented for each click
        /// </summary>
        public int Increment
        {
            get
            {
                return this._ctrlProp.GetXmlNodeInt("@inc", 1);
            }
            set
            {
                if(value < 0 || value >3000)
                {
                    throw (new ArgumentOutOfRangeException("Increment must be between 0 and 3000"));
                }

                this._ctrlProp.SetXmlNodeInt("@inc", value);
                this._vmlProp.SetXmlNodeInt("x:Inc", value);
            }
        }
        /// <summary>
        /// The number of items to move the scroll bar on a page click. Null is default
        /// </summary>
        public int? Page
        {
            get
            {
                return this._ctrlProp.GetXmlNodeIntNull("@page");
            }
            set
            {
                if (value.HasValue && (value < 0 || value > 3000))
                {
                    throw (new ArgumentOutOfRangeException("Page must be between 0 and 3000"));
                }

                this._ctrlProp.SetXmlNodeInt("@page", value);
                this._vmlProp.SetXmlNodeInt("x:Page", value);
            }
        }
        /// <summary>
        /// The value when a scroll bar is at it's minimum
        /// </summary>
        public int MinValue
        {
            get
            {
                return this._ctrlProp.GetXmlNodeInt("@min", 0);
            }
            set
            {
                if (value < 0 || value > 30000)
                {
                    throw (new ArgumentOutOfRangeException("MinValue must be between 0 and 3000"));
                }

                this._ctrlProp.SetXmlNodeInt("@min", value);
            }
        }
        /// <summary>
        /// The value when a scroll bar is at it's maximum
        /// </summary>
        public int MaxValue
        {
            get
            {
                return this._ctrlProp.GetXmlNodeInt("@max", 30000);
            }
            set
            {
                if (value < 0 || value > 30000)
                {
                    throw (new ArgumentOutOfRangeException("MaxValue must be between 0 and 30000"));
                }

                this._ctrlProp.SetXmlNodeInt("@max", value);
            }
        }
        /// <summary>
        /// The value of the scroll bar.
        /// </summary>
        public int Value
        {
            get
            {
                return this._ctrlProp.GetXmlNodeInt("@val", 0);
            }
            set
            {
                if (value < 0 || value > 30000)
                {
                    throw (new ArgumentOutOfRangeException("Value must be between 0 and 30000"));
                }

                this._ctrlProp.SetXmlNodeInt("@val", value);
                this._vmlProp.SetXmlNodeInt("x:Val", value);

                this.SetLinkedCellValue(value);
            }
        }
    }
}