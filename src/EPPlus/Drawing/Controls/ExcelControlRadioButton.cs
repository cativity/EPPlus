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
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    /// <summary>
    /// Represents a radio button form control
    /// </summary>
    public class ExcelControlRadioButton : ExcelControlWithColorsAndLines
    {
        internal ExcelControlRadioButton(ExcelDrawings drawings, XmlNode drawNode, ControlInternal control, ZipPackagePart part, XmlDocument controlPropertiesXml, ExcelGroupShape parent = null)
            : base(drawings, drawNode, control, part, controlPropertiesXml, parent)            
        {
        }
        internal ExcelControlRadioButton(ExcelDrawings drawings, XmlElement drawNode, string name, ExcelGroupShape parent=null) : base(drawings, drawNode, name, parent)
        {
        }

        /// <summary>
        /// The type of form control
        /// </summary>
        public override eControlType ControlType => eControlType.RadioButton;
        /// <summary>
        /// Gets or sets the state of the radio box.
        /// </summary>
        public bool Checked
        {
            get
            {
                return this._ctrlProp.GetXmlNodeString("@checked")=="Checked";
            }
            set
            {
                this._ctrlProp.SetXmlNodeString("@checked", value?"Checked":"Unchecked");
            }
        }
        /// <summary>
        /// Gets or sets the address to the cell that is linked to the control. 
        /// </summary>
        public new ExcelAddressBase LinkedCell
        {
            get
            {
                ExcelAddressBase? v= this.LinkedGroup;
                if(v!=null)
                {
                    return v;
                }
                return this.FmlaLink;
            }
            set
            {
                if (this.LinkedGroup == null)
                {
                    this.FmlaLink = value;
                }
                else
                {
                    this.LinkedGroup = value;
                }
            }
        }
        /// <summary>
        /// Gets or sets if the radio button is the first button in a set of radio buttons
        /// </summary>
        public bool FirstButton
        {
            get
            {
                return this._ctrlProp.GetXmlNodeBool("@firstButton");
            }
            set
            {
                this._ctrlProp.SetXmlNodeBool("@firstButton", value);
                this._vmlProp.SetBoolNode("x:FirstButton", value);
            }
        }
    }
}