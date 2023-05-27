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
using System.Linq;
using System.Text;
using System.Drawing;
using System.Xml;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.Style.Dxf;

/// <summary>
/// A base class for differential formatting font styles
/// </summary>
public class ExcelDxfFontBase : DxfStyleBase
{
    internal ExcelDxfFontBase(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
        : base(styles, callback)
    {
        this.Color = new ExcelDxfColor(styles, eStyleClass.Font, callback);
    }

    bool? _bold;

    /// <summary>
    /// Font bold
    /// </summary>
    public bool? Bold
    {
        get { return this._bold; }
        set
        {
            this._bold = value;
            this._callback?.Invoke(eStyleClass.Font, eStyleProperty.Bold, value);
        }
    }

    bool? _italic;

    /// <summary>
    /// Font Italic
    /// </summary>
    public bool? Italic
    {
        get { return this._italic; }
        set
        {
            this._italic = value;
            this._callback?.Invoke(eStyleClass.Font, eStyleProperty.Italic, value);
        }
    }

    bool? _strike;

    /// <summary>
    /// Font-Strikeout
    /// </summary>
    public bool? Strike
    {
        get { return this._strike; }
        set
        {
            this._strike = value;
            this._callback?.Invoke(eStyleClass.Font, eStyleProperty.Strike, value);
        }
    }

    /// <summary>
    /// The color of the text
    /// </summary>
    public ExcelDxfColor Color { get; set; }

    ExcelUnderLineType? _underline;

    /// <summary>
    /// The underline type
    /// </summary>
    public ExcelUnderLineType? Underline
    {
        get { return this._underline; }
        set
        {
            this._underline = value;
            this._callback?.Invoke(eStyleClass.Font, eStyleProperty.UnderlineType, value);
        }
    }

    /// <summary>
    /// The id
    /// </summary>
    internal override string Id
    {
        get
        {
            return GetAsString(this.Bold)
                   + "|"
                   + GetAsString(this.Italic)
                   + "|"
                   + GetAsString(this.Strike)
                   + "|"
                   + (this.Color == null ? "" : this.Color.Id)
                   + "|"
                   + GetAsString(this.Underline)
                   + "|||||||||";
        }
    }

    /// <summary>
    /// Creates the the xml node
    /// </summary>
    /// <param name="helper">The xml helper</param>
    /// <param name="path">The X Path</param>
    internal override void CreateNodes(XmlHelper helper, string path)
    {
        helper.CreateNode(path);
        SetValueBool(helper, path + "/d:b/@val", this.Bold);
        SetValueBool(helper, path + "/d:i/@val", this.Italic);
        SetValueBool(helper, path + "/d:strike/@val", this.Strike);
        SetValue(helper, path + "/d:u/@val", this.Underline == null ? null : this.Underline.ToEnumString());
        SetValueColor(helper, path + "/d:color", this.Color);
    }

    internal override void SetStyle()
    {
        if (this._callback != null)
        {
            this._callback.Invoke(eStyleClass.Font, eStyleProperty.Bold, this._bold);
            this._callback.Invoke(eStyleClass.Font, eStyleProperty.Italic, this._italic);
            this._callback.Invoke(eStyleClass.Font, eStyleProperty.Strike, this._strike);
            this._callback.Invoke(eStyleClass.Font, eStyleProperty.UnderlineType, this.Underline);
            this.Color.SetStyle();
        }
    }

    /// <summary>
    /// If the object has any properties set
    /// </summary>
    public override bool HasValue
    {
        get { return this.Bold != null || this.Italic != null || this.Strike != null || this.Underline != null || this.Color.HasValue; }
    }

    /// <summary>
    /// Clears all properties
    /// </summary>
    public override void Clear()
    {
        this.Bold = null;
        this.Italic = null;
        this.Strike = null;
        this.Underline = null;
        this.Color.Clear();
    }

    /// <summary>
    /// Clone the object
    /// </summary>
    /// <returns>A new instance of the object</returns>
    internal override DxfStyleBase Clone()
    {
        return new ExcelDxfFontBase(this._styles, this._callback)
        {
            Bold = this.Bold, Color = (ExcelDxfColor)this.Color.Clone(), Italic = this.Italic, Strike = this.Strike, Underline = this.Underline
        };
    }

    internal override void SetValuesFromXml(XmlHelper helper)
    {
        if (helper.ExistsNode("d:font"))
        {
            this.Bold = helper.GetXmlNodeBoolNullableWithVal("d:font/d:b");
            this.Italic = helper.GetXmlNodeBoolNullableWithVal("d:font/d:i");
            this.Strike = helper.GetXmlNodeBoolNullableWithVal("d:font/d:strike");
            this.Underline = GetUnderLine(helper);
            this.Color = this.GetColor(helper, "d:font/d:color", eStyleClass.Font);
        }
    }

    private static ExcelUnderLineType? GetUnderLine(XmlHelper helper)
    {
        if (helper.ExistsNode("d:font/d:u"))
        {
            string? v = helper.GetXmlNodeString("d:font/d:u/@val");

            if (string.IsNullOrEmpty(v))
            {
                return ExcelUnderLineType.Single;
            }
            else
            {
                return GetUnderLineEnum(v);
            }
        }
        else
        {
            return null;
        }
    }
}