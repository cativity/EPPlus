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
using System.Xml;

namespace OfficeOpenXml.Style.Dxf;

/// <summary>
/// A single border line of a drawing in a differential formatting record
/// </summary>
public class ExcelDxfBorderItem : DxfStyleBase
{
    eStyleClass _styleClass;

    internal ExcelDxfBorderItem(ExcelStyles styles, eStyleClass styleClass, Action<eStyleClass, eStyleProperty, object> callback)
        : base(styles, callback)
    {
        this._styleClass = styleClass;
        this.Color = new ExcelDxfColor(styles, this._styleClass, callback);
    }

    ExcelBorderStyle? _style;

    /// <summary>
    /// The border style
    /// </summary>
    public ExcelBorderStyle? Style
    {
        get => this._style;
        set
        {
            this._style = value;
            this._callback?.Invoke(this._styleClass, eStyleProperty.Style, value);
        }
    }

    /// <summary>
    /// The color of the border
    /// </summary>
    public ExcelDxfColor Color { get; internal set; }

    /// <summary>
    /// The Id
    /// </summary>
    internal override string Id => GetAsString(this.Style) + "|" + (this.Color == null ? "" : this.Color.Id);

    /// <summary>
    /// Creates the the xml node
    /// </summary>
    /// <param name="helper">The xml helper</param>
    /// <param name="path">The X Path</param>
    internal override void CreateNodes(XmlHelper helper, string path)
    {
        SetValueEnum(helper, path + "/@style", this.Style);
        SetValueColor(helper, path + "/d:color", this.Color);
    }

    /// <summary>
    /// If the object has any properties set
    /// </summary>
    public override bool HasValue => (this.Style != null && this.Style != ExcelBorderStyle.None) || this.Color.HasValue;

    /// <summary>
    /// Clears all properties
    /// </summary>
    public override void Clear()
    {
        this.Style = null;
        this.Color.Clear();
    }

    internal override void SetStyle()
    {
        if (this._callback != null)
        {
            this._callback.Invoke(this._styleClass, eStyleProperty.Style, this._style);
            this.Color.SetStyle();
        }
    }

    /// <summary>
    /// Clone the object
    /// </summary>
    /// <returns>A new instance of the object</returns>
    internal override DxfStyleBase Clone() => new ExcelDxfBorderItem(this._styles, this._styleClass, this._callback) { Style = this.Style, Color = this.Color };
}