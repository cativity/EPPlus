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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Controls;

/// <summary>
/// Represents a list box form control.
/// </summary>
public class ExcelControlListBox : ExcelControlList
{
    internal ExcelControlListBox(ExcelDrawings drawings, XmlElement drawNode, string name, ExcelGroupShape parent=null) : base(drawings, drawNode, name, parent)
    {
        this.SetSize(150, 100); //Default size
    }
    internal ExcelControlListBox(ExcelDrawings drawings, XmlNode drawNode, ControlInternal control, ZipPackagePart part, XmlDocument controlPropertiesXml, ExcelGroupShape parent = null)
        : base(drawings, drawNode, control, part, controlPropertiesXml, parent)
    {
    }

    /// <summary>
    /// The type of form control
    /// </summary>
    public override eControlType ControlType => eControlType.ListBox;
    /// <summary>
    /// The type of selection
    /// </summary>
    public eSelectionType SelectionType
    {
        get
        {
            return this._ctrlProp.GetXmlNodeString("@seltype").ToEnum(eSelectionType.Single);
        }
        set
        {
            this._ctrlProp.SetXmlNodeString("@seltype", value.ToEnumString());
            this._vmlProp.SetXmlNodeString("x:SelType", value.ToString());
        }
    }
    /// <summary>
    /// If <see cref="SelectionType"/> is Multi or Extended this array contains the selected indicies. Index is zero based. 
    /// </summary>
    public int[] MultiSelection
    {
        get
        {
            string? s = this._ctrlProp.GetXmlNodeString("@multiSel");
            if (string.IsNullOrEmpty(s))
            {
                return null;
            }
            else
            {
                string[]? a = s.Split(',');
                try
                {
                    return a.Select(x => int.Parse(x) - 1).ToArray();
                }
                catch
                {
                    return null;
                }
            }
        }
        set
        {
            if (value == null)
            {
                this._ctrlProp.DeleteNode("@multiSel");
                this._vmlProp.DeleteNode("x:MultiSel");
            }
            string? v = value.Select(x => (x + 1).ToString(CultureInfo.InvariantCulture)).Aggregate((x, y) => x + "," + y);
            this._ctrlProp.SetXmlNodeString("selType", v);
            this._vmlProp.SetXmlNodeString("x:MultiSel", v);
        }
    }
    internal override void UpdateXml()
    {
        base.UpdateXml();
        ((ExcelControlList)this).Page = (int)Math.Round((this._height / 14));
    }
}