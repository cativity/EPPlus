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
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// A page / report filter field
/// </summary>
public class ExcelPivotTablePageFieldSettings : XmlHelper
{
    //internal ExcelPivotTableField _field;

    internal ExcelPivotTablePageFieldSettings(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTableField field, int index)
        : base(ns, topNode)
    {
        if (this.GetXmlNodeString("@hier") == "")
        {
            this.Hier = -1;
        }

        //this._field = field;
    }

    internal int Index
    {
        get => this.GetXmlNodeInt("@fld");
        set => this.SetXmlNodeString("@fld", value.ToString());
    }

    /// <summary>
    /// The Name of the field
    /// </summary>
    public string Name
    {
        get => this.GetXmlNodeString("@name");
        set => this.SetXmlNodeString("@name", value);
    }

    /***** Dont work. Need items to be populated. ****/
    /// <summary>
    /// The selected item 
    /// </summary>
    internal int SelectedItem
    {
        get => this.GetXmlNodeInt("@item");
        set
        {
            if (value < 0)
            {
                this.DeleteNode("@item");
            }
            else
            {
                this.SetXmlNodeString("@item", value.ToString());
            }
        }
    }

    internal int NumFmtId
    {
        get => this.GetXmlNodeInt("@numFmtId");
        set => this.SetXmlNodeString("@numFmtId", value.ToString());
    }

    internal int Hier
    {
        get => this.GetXmlNodeInt("@hier");
        set => this.SetXmlNodeString("@hier", value.ToString());
    }
}