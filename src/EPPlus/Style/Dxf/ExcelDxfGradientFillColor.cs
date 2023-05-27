/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/29/2021         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/

using System;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf;

/// <summary>
/// Represents a position of a color in a gradient list for differencial styles.
/// </summary>
public class ExcelDxfGradientFillColor : DxfStyleBase
{
    internal ExcelDxfGradientFillColor(ExcelStyles styles, double position, Action<eStyleClass, eStyleProperty, object> callback)
        : base(styles, callback)
    {
        this.Position = position;
        eStyleClass styleClass = position == 0 ? eStyleClass.FillGradientColor1 : eStyleClass.FillGradientColor2;
        this.Color = new ExcelDxfColor(styles, styleClass, callback);
    }

    /// <summary>
    /// The position of the color 
    /// </summary>
    public double Position { get; }

    /// <summary>
    /// The color to use at the position
    /// </summary>
    public ExcelDxfColor Color { get; internal set; }

    /// <summary>
    /// If the object has any properties set
    /// </summary>
    public override bool HasValue
    {
        get { return this.Color.HasValue; }
    }

    internal override string Id
    {
        get { return this.Position.ToString() + "|" + this.Color.Id; }
    }

    /// <summary>
    /// Clears all colors
    /// </summary>
    public override void Clear()
    {
        this.Color.Clear();
    }

    internal override DxfStyleBase Clone()
    {
        return new ExcelDxfGradientFillColor(this._styles, this.Position, this._callback) { Color = (ExcelDxfColor)this.Color.Clone() };
    }

    internal override void CreateNodes(XmlHelper helper, string path)
    {
        XmlNode? node = helper.CreateNode(path + "d:stop", false, true);
        XmlHelper? stopHelper = XmlHelperFactory.Create(helper.NameSpaceManager, node);
        SetValue(stopHelper, "@position", this.Position / 100);
        SetValueColor(stopHelper, "d:color", this.Color);
    }

    internal override void SetStyle()
    {
        this.Color.SetStyle();
    }
}