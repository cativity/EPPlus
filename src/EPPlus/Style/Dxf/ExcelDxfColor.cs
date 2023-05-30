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
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.Style.Dxf;

/// <summary>
/// A color in a differential formatting record
/// </summary>
public class ExcelDxfColor : DxfStyleBase

{
    eStyleClass _styleClass;

    internal ExcelDxfColor(ExcelStyles styles, eStyleClass styleClass, Action<eStyleClass, eStyleProperty, object> callback)
        : base(styles, callback)
    {
        this._styleClass = styleClass;
    }

    eThemeSchemeColor? _theme;

    /// <summary>
    /// Gets or sets a theme color
    /// </summary>
    public eThemeSchemeColor? Theme
    {
        get { return this._theme; }
        set
        {
            this._theme = value;
            this._callback?.Invoke(this._styleClass, eStyleProperty.Theme, value);
        }
    }

    int? _index;

    /// <summary>
    /// Gets or sets an indexed color
    /// </summary>
    public int? Index
    {
        get { return this._index; }
        set
        {
            this._index = value;
            this._callback?.Invoke(this._styleClass, eStyleProperty.IndexedColor, value);
        }
    }

    bool? _auto;

    /// <summary>
    /// Gets or sets the color to automativ
    /// </summary>
    public bool? Auto
    {
        get { return this._auto; }
        set
        {
            this._auto = value;
            this._callback?.Invoke(this._styleClass, eStyleProperty.AutoColor, value);
        }
    }

    double? _tint;

    /// <summary>
    /// Gets or sets the Tint value for the color
    /// </summary>
    public double? Tint
    {
        get { return this._tint; }
        set
        {
            this._tint = value;
            this._callback?.Invoke(this._styleClass, eStyleProperty.Tint, value);
        }
    }

    Color? _color;

    /// <summary>
    /// Sets the color.
    /// </summary>
    public Color? Color
    {
        get { return this._color; }
        set
        {
            this._color = value;
            this._callback?.Invoke(this._styleClass, eStyleProperty.Color, value);
        }
    }

    /// <summary>
    /// The Id
    /// </summary>
    internal override string Id
    {
        get
        {
            return GetAsString(this.Theme)
                   + "|"
                   + GetAsString(this.Index)
                   + "|"
                   + GetAsString(this.Auto)
                   + "|"
                   + GetAsString(this.Tint)
                   + "|"
                   + GetAsString(this.Color == null ? "" : this.Color.Value.ToArgb().ToString("x"));
        }
    }

    /// <summary>
    /// Set the color of the drawing
    /// </summary>
    /// <param name="color">The color</param>
    public void SetColor(Color color)
    {
        this.Theme = null;
        this.Auto = null;
        this.Index = null;
        this.Color = color;
    }

    /// <summary>
    /// Set the color of the drawing
    /// </summary>
    /// <param name="color">The color</param>
    public void SetColor(eThemeSchemeColor color)
    {
        this.Color = null;
        this.Auto = null;
        this.Index = null;
        this.Theme = color;
    }

    /// <summary>
    /// Set the color of the drawing
    /// </summary>
    /// <param name="color">The color</param>
    public void SetColor(ExcelIndexedColor color)
    {
        this.Color = null;
        this.Theme = null;
        this.Auto = null;
        this.Index = (int)color;
    }

    /// <summary>
    /// Set the color to automatic
    /// </summary>
    public void SetAuto()
    {
        this.Color = null;
        this.Theme = null;
        this.Index = null;
        this.Auto = true;
    }

    internal override void SetStyle()
    {
        if (this._callback != null)
        {
            this._callback.Invoke(this._styleClass, eStyleProperty.Color, this._color);
            this._callback.Invoke(this._styleClass, eStyleProperty.Theme, this._theme);
            this._callback.Invoke(this._styleClass, eStyleProperty.IndexedColor, this._index);
            this._callback.Invoke(this._styleClass, eStyleProperty.AutoColor, this._auto);
            this._callback.Invoke(this._styleClass, eStyleProperty.Tint, this._tint);
        }
    }

    /// <summary>
    /// Clone the object
    /// </summary>
    /// <returns>A new instance of the object</returns>
    internal override DxfStyleBase Clone()
    {
        return new ExcelDxfColor(this._styles, this._styleClass, this._callback)
        {
            Theme = this.Theme, Index = this.Index, Color = this.Color, Auto = this.Auto, Tint = this.Tint
        };
    }

    /// <summary>
    /// If the object has any properties set
    /// </summary>
    public override bool HasValue
    {
        get { return this.Theme != null || this.Index != null || this.Auto != null || this.Tint != null || this.Color != null; }
    }

    /// <summary>
    /// Clears all properties
    /// </summary>
    public override void Clear()
    {
        this.Theme = null;
        this.Index = null;
        this.Auto = null;
        this.Tint = null;
        this.Color = null;
    }

    /// <summary>
    /// Creates the the xml node
    /// </summary>
    /// <param name="helper">The xml helper</param>
    /// <param name="path">The X Path</param>
    internal override void CreateNodes(XmlHelper helper, string path)
    {
        throw new NotImplementedException();
    }
}