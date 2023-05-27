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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable;

/// <summary>
/// A pivot table field Item. Used for grouping.
/// </summary>
public class ExcelPivotTableFieldItem
{
    [Flags]
    internal enum eBoolFlags
    {
        Hidden=0x1,
        ShowDetails = 0x2,
        C = 0x4,
        D = 0x8,
        E = 0x10,
        F = 0x20,
        M = 0x40,
        S = 0x80
    }
    internal eBoolFlags flags=eBoolFlags.ShowDetails|eBoolFlags.E;
    internal ExcelPivotTableFieldItem()
    {
    }
    internal ExcelPivotTableFieldItem (XmlElement node)
    {
        foreach(XmlAttribute a in node.Attributes)
        {
            switch(a.LocalName)
            {
                case "c":
                    this.C = XmlHelper.GetBoolFromString(a.Value);
                    break;
                case "d":
                    this.D = XmlHelper.GetBoolFromString(a.Value);
                    break;
                case "e":
                    this.E = XmlHelper.GetBoolFromString(a.Value);
                    break;
                case "f":
                    this.F = XmlHelper.GetBoolFromString(a.Value);
                    break;
                case "h":
                    this.Hidden = XmlHelper.GetBoolFromString(a.Value);
                    break;
                case "m":
                    this.M = XmlHelper.GetBoolFromString(a.Value);
                    break;
                case "n":
                    this.Text = a.Value;
                    break;
                case "s":
                    this.S = XmlHelper.GetBoolFromString(a.Value);
                    break;
                case "sd":
                    this.ShowDetails = XmlHelper.GetBoolFromString(a.Value);
                    break;
                case "t":
                    this.Type = a.Value.ToEnum(eItemType.Data);
                    break;
                case "x":
                    this.X = int.Parse(a.Value);
                    break;
            }
        }
    }
    /// <summary>
    /// The custom text of the item. Unique values only
    /// </summary>
    public string Text { get; set; }
    /// <summary>
    /// The value of the item
    /// </summary>
    public object Value { get; internal set; }
    /// <summary>
    /// A flag indicating if the items are hidden
    /// </summary>
    public bool Hidden 
    { 
        get
        {
            return (this.flags & eBoolFlags.Hidden) == eBoolFlags.Hidden;
        }
        set
        {
            if (this.Type != eItemType.Data)
            {
                throw new InvalidOperationException("Hidden can only be set for items of type Data");
            }

            this.SetFlag(eBoolFlags.Hidden, value);
        }
    }

    /// <summary>
    /// A flag indicating if the items expanded or collapsed.
    /// </summary>
    public bool ShowDetails
    {
        get
        {
            return (this.flags & eBoolFlags.ShowDetails) == eBoolFlags.ShowDetails;
        }
        set
        {
            this.SetFlag(eBoolFlags.ShowDetails, value);
        }
    }
    internal bool C
    {
        get
        {
            return (this.flags & eBoolFlags.C) == eBoolFlags.C;
        }
        set
        {
            this.SetFlag(eBoolFlags.C, value);
        }
    }
    internal bool D
    {
        get
        {
            return (this.flags & eBoolFlags.D) == eBoolFlags.D;
        }
        set
        {
            this.SetFlag(eBoolFlags.D, value);
        }
    }
    internal bool E
    {
        get
        {
            return (this.flags & eBoolFlags.E) == eBoolFlags.E;
        }
        set
        {
            this.SetFlag(eBoolFlags.E, value);
        }
    }
    internal bool F
    {
        get
        {
            return (this.flags & eBoolFlags.F) == eBoolFlags.F;
        }
        set
        {
            this.SetFlag(eBoolFlags.F, value);
        }
    }
    internal bool M
    {
        get
        {
            return (this.flags & eBoolFlags.M) == eBoolFlags.M;
        }
        set
        {
            this.SetFlag(eBoolFlags.M, value);
        }
    }
    internal bool S
    {
        get
        {
            return (this.flags & eBoolFlags.S) == eBoolFlags.S;
        }
        set
        {
            this.SetFlag(eBoolFlags.S, value);
        }
    }
    internal int X { get; set; } = -1;
    internal eItemType Type { get; set; }

    internal void GetXmlString(StringBuilder sb)
    {
        if (this.X == -1 && this.Type == eItemType.Data)
        {
            return;
        }

        sb.Append("<item");
        if(this.X>-1)
        {
            sb.AppendFormat(" x=\"{0}\"", this.X);
        }
        if(this.Type!=eItemType.Data)
        {
            sb.AppendFormat(" t=\"{0}\"", this.Type.ToEnumString());
        }
        if(!string.IsNullOrEmpty(this.Text))
        {
            sb.AppendFormat(" n=\"{0}\"", Utils.ConvertUtil.ExcelEscapeString(this.Text));
        }
        AddBool(sb,"h", this.Hidden);
        AddBool(sb, "sd", this.ShowDetails, true);
        AddBool(sb, "c", this.C);
        AddBool(sb, "d", this.D);
        AddBool(sb, "e", this.E, true);
        AddBool(sb, "f", this.F);
        AddBool(sb, "m", this.M);
        AddBool(sb, "s", this.S);
        sb.Append("/>");
    }

    private static void AddBool(StringBuilder sb, string attrName, bool b, bool defaultValue=false)
    {
        if(b != defaultValue)
        {
            sb.AppendFormat(" {0}=\"{1}\"",attrName, b?"1":"0");
        }
    }
    private void SetFlag(eBoolFlags flag, bool value)
    {
        if(value)
        {
            this.flags |= flag;
        }
        else
        {
            this.flags &= ~flag;
        }
    }
}