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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls
{
    /// <summary>
    /// An abstract class used by form control list objects
    /// </summary>
    public abstract class ExcelControlList : ExcelControl
    {
        internal ExcelControlList(ExcelDrawings drawings, XmlElement drawNode, string name, ExcelGroupShape parent=null) : base(drawings, drawNode, name, parent)
        {
        }

        internal ExcelControlList(ExcelDrawings drawings, XmlNode drawNode, ControlInternal control, ZipPackagePart part, XmlDocument controlPropertiesXml, ExcelGroupShape parent = null)
            : base(drawings, drawNode, control, part, controlPropertiesXml, parent)
        {
        }
        /// <summary>
        /// The range to the items populating the list.
        /// </summary>
        public ExcelAddressBase InputRange 
        { 
            get
            {
                string? range = this._ctrlProp.GetXmlNodeString("@fmlaRange");
                if(ExcelCellBase.IsValidAddress(range))
                {
                    return new ExcelAddressBase(range);
                }
                return null;
            }
            set
            {
                if (value == null)
                {
                    this._ctrlProp.DeleteNode("@fmlaRange");
                    this._vmlProp.DeleteNode("x:FmlaRange");
                }
                if (value.WorkSheetName.Equals(this._drawings.Worksheet.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    this._ctrlProp.SetXmlNodeString("@fmlaRange", value.Address);
                    this._vmlProp.SetXmlNodeString("x:FmlaRange", value.Address);
                }
                else
                {
                    this._ctrlProp.SetXmlNodeString("@fmlaRange", value.FullAddress);
                    this._vmlProp.SetXmlNodeString("x:FmlaRange", value.FullAddress);
                }
            }
        }
        /// <summary>
        /// The index of a selected item in the list. 
        /// </summary>
        public int SelectedIndex
        {
            get
            {
                return this._ctrlProp.GetXmlNodeInt("@sel", 0) - 1;
            }
            set
            {
                if (value <= 0)
                {
                    this._ctrlProp.DeleteNode("@sel");
                    this._vmlProp.DeleteNode("x:Sel");
                }
                else
                {
                    this._ctrlProp.SetXmlNodeInt("@sel", value);
                    this._vmlProp.SetXmlNodeInt("x:Sel", value);
                }
            }
        }
        internal int Page
        {
            get
            {
                return this._vmlProp.GetXmlNodeInt("x:Page");
            }
            set
            {
                this._ctrlProp.SetXmlNodeInt("@page", value);
                this._vmlProp.SetXmlNodeInt("x:Page", value);
            }
        }
    }
}