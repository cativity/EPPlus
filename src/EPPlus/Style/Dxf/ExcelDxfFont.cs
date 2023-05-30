/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/01/2021         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/

using System;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.Style.Dxf;

/// <summary>
/// A font in a differential formatting record
/// </summary>
public class ExcelDxfFont : ExcelDxfFontBase
{
    internal ExcelDxfFont(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
        : base(styles, callback)
    {
    }

    float? _size;

    /// <summary>
    /// The font size 
    /// </summary>
    public float? Size
    {
        get => this._size;
        set
        {
            this._size = value;
            this._callback?.Invoke(eStyleClass.Font, eStyleProperty.Size, value);
        }
    }

    string _name;

    /// <summary>
    /// The name of the font
    /// </summary>
    public string Name
    {
        get => this._name;
        set
        {
            this._name = value;
            this._callback?.Invoke(eStyleClass.Font, eStyleProperty.Name, value);
        }
    }

    int? _family;

    /// <summary>
    /// Font family 
    /// </summary>
    public int? Family
    {
        get => this._family;
        set
        {
            this._family = value;
            this._callback?.Invoke(eStyleClass.Font, eStyleProperty.Family, value);
        }
    }

    ExcelVerticalAlignmentFont _verticalAlign = ExcelVerticalAlignmentFont.None;

    /// <summary>
    /// Font-Vertical Align
    /// </summary>
    public ExcelVerticalAlignmentFont VerticalAlign
    {
        get => this._verticalAlign;
        set
        {
            this._verticalAlign = value;
            this._callback?.Invoke(eStyleClass.Font, eStyleProperty.VerticalAlign, value);
        }
    }

    bool? _outline;

    /// <summary>
    /// Displays only the inner and outer borders of each character. Similar to bold
    /// </summary>
    public bool? Outline
    {
        get => this._outline;
        set => this._outline = value;
    }

    bool? _shadow;

    /// <summary>
    /// Shadow for the font. Used on Macintosh only.
    /// </summary>
    public bool? Shadow
    {
        get => this._shadow;
        set => this._shadow = value;
    }

    bool? _condense;

    /// <summary>
    /// Condence (squeeze it together). Used on Macintosh only.
    /// </summary>
    public bool? Condense
    {
        get => this._condense;
        set => this._condense = value;
    }

    bool? _extend;

    /// <summary>
    /// Extends or stretches the text. Legacy property used in older speadsheet applications.
    /// </summary>
    public bool? Extend
    {
        get => this._extend;
        set => this._extend = value;
    }

    eThemeFontCollectionType? _scheme;

    /// <summary>
    /// Which font scheme to use from the theme
    /// </summary>
    public eThemeFontCollectionType? Scheme
    {
        get => this._scheme;
        set
        {
            this._scheme = value;
            this._callback?.Invoke(eStyleClass.Font, eStyleProperty.Scheme, value);
        }
    }

    /// <summary>
    /// The Id to identify the font uniquely
    /// </summary>
    internal override string Id =>
        GetAsString(this.Bold)
        + "|"
        + GetAsString(this.Italic)
        + "|"
        + GetAsString(this.Strike)
        + "|"
        + (this.Color == null ? "" : this.Color.Id)
        + "|"
        + GetAsString(this.Underline)
        + "|"
        + GetAsString(this.Name)
        + "|"
        + GetAsString(this.Size)
        + "|"
        + GetAsString(this.Family)
        + "|"
        + this.GetVAlign()
        + "|"
        + GetAsString(this.Outline)
        + "|"
        + GetAsString(this.Shadow)
        + "|"
        + GetAsString(this.Condense)
        + "|"
        + GetAsString(this.Extend)
        + "|"
        + GetAsString(this.Scheme);

    private string GetVAlign() => this.VerticalAlign == ExcelVerticalAlignmentFont.None ? "" : GetAsString(this.VerticalAlign);

    /// <summary>
    /// Clone the object
    /// </summary>
    /// <returns>A new instance of the object</returns>
    internal override DxfStyleBase Clone() =>
        new ExcelDxfFont(this._styles, this._callback)
        {
            Name = this.Name,
            Size = this.Size,
            Family = this.Family,
            Bold = this.Bold,
            Color = (ExcelDxfColor)this.Color.Clone(),
            Italic = this.Italic,
            Strike = this.Strike,
            Underline = this.Underline,
            Condense = this.Condense,
            Extend = this.Extend,
            Scheme = this.Scheme,
            Outline = this.Outline,
            Shadow = this.Shadow,
            VerticalAlign = this.VerticalAlign
        };

    /// <summary>
    /// If the object has any properties set
    /// </summary>
    public override bool HasValue =>
        base.HasValue
        || string.IsNullOrEmpty(this.Name) == false
        || this.Size.HasValue
        || this.Family.HasValue
        || this.Condense.HasValue
        || this.Extend.HasValue
        || this.Scheme.HasValue
        || this.Outline.HasValue
        || this.Shadow.HasValue
        || this.VerticalAlign != ExcelVerticalAlignmentFont.None;

    /// <summary>
    /// Clears all properties
    /// </summary>
    public override void Clear()
    {
        base.Clear();
        this.Name = null;
        this.Size = null;
        this.Family = null;
        this.Condense = null;
        this.Extend = null;
        this.Scheme = null;
        this.Outline = null;
        this.Shadow = null;
        this.VerticalAlign = ExcelVerticalAlignmentFont.None;
    }

    internal override void CreateNodes(XmlHelper helper, string path)
    {
        _ = helper.CreateNode(path);
        SetValueBool(helper, path + "/d:b/@val", this.Bold);
        SetValueBool(helper, path + "/d:i/@val", this.Italic);
        SetValueBool(helper, path + "/d:strike/@val", this.Strike);
        SetValue(helper, path + "/d:u/@val", this.Underline == null ? null : this.Underline.ToEnumString());
        SetValueBool(helper, path + "/d:condense/@val", this.Condense);
        SetValueBool(helper, path + "/d:extend/@val", this.Extend);
        SetValueBool(helper, path + "/d:outline/@val", this.Outline);
        SetValueBool(helper, path + "/d:shadow/@val", this.Shadow);
        SetValue(helper, path + "/d:sz/@val", this.Size);
        SetValueColor(helper, path + "/d:color", this.Color);
        SetValue(helper, path + "/d:name/@val", this.Name);
        SetValue(helper, path + "/d:family/@val", this.Family);
        SetValue(helper, path + "/d:vertAlign/@val", this.VerticalAlign == ExcelVerticalAlignmentFont.None ? null : this.VerticalAlign.ToEnumString());
    }

    internal override void SetValuesFromXml(XmlHelper helper)
    {
        this.Size = helper.GetXmlNodeIntNull("d:font/d:sz/@val");

        base.SetValuesFromXml(helper);

        this.Name = helper.GetXmlNodeString("d:font/d:name/@val");
        this.Condense = helper.GetXmlNodeBoolNullable("d:font/d:condense/@val");
        this.Extend = helper.GetXmlNodeBoolNullable("d:font/d:extend/@val");
        this.Outline = helper.GetXmlNodeBoolNullable("d:font/d:outline/@val");

        string? v = helper.GetXmlNodeString("d:font/d:vertAlign/@val");
        this.VerticalAlign = string.IsNullOrEmpty(v) ? ExcelVerticalAlignmentFont.None : v.ToEnum(ExcelVerticalAlignmentFont.None);

        this.Family = helper.GetXmlNodeIntNull("d:font/d:family/@val");
        this.Scheme = helper.GetXmlEnumNull<eThemeFontCollectionType>("d:font/d:scheme/@val");
        this.Shadow = helper.GetXmlNodeBoolNullable("d:font/d:shadow/@val");
    }

    internal override void SetStyle()
    {
        if (this._callback != null)
        {
            base.SetStyle();
            this._callback?.Invoke(eStyleClass.Font, eStyleProperty.Name, this._name);
            this._callback?.Invoke(eStyleClass.Font, eStyleProperty.Size, this._size);
            this._callback?.Invoke(eStyleClass.Font, eStyleProperty.Family, this._family);
            this._callback?.Invoke(eStyleClass.Font, eStyleProperty.Scheme, this._scheme);
            this._callback?.Invoke(eStyleClass.Font, eStyleProperty.VerticalAlign, this._verticalAlign);
        }
    }
}