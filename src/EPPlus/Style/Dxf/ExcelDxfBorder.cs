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
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf;

/// <summary>
/// The border style of a drawing in a differential formatting record
/// </summary>
public class ExcelDxfBorderBase : DxfStyleBase
{
    internal ExcelDxfBorderBase(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
        : base(styles, callback)
    {
        this.Left = new ExcelDxfBorderItem(this._styles, eStyleClass.BorderLeft, callback);
        this.Right = new ExcelDxfBorderItem(this._styles, eStyleClass.BorderRight, callback);
        this.Top = new ExcelDxfBorderItem(this._styles, eStyleClass.BorderTop, callback);
        this.Bottom = new ExcelDxfBorderItem(this._styles, eStyleClass.BorderBottom, callback);
        this.Vertical = new ExcelDxfBorderItem(this._styles, eStyleClass.Border, callback);
        this.Horizontal = new ExcelDxfBorderItem(this._styles, eStyleClass.Border, callback);
    }
    /// <summary>
    /// Left border style
    /// </summary>
    public ExcelDxfBorderItem Left
    {
        get;
        internal set;
    }
    /// <summary>
    /// Right border style
    /// </summary>
    public ExcelDxfBorderItem Right
    {
        get;
        internal set;
    }
    /// <summary>
    /// Top border style
    /// </summary>
    public ExcelDxfBorderItem Top
    {
        get;
        internal set;
    }
    /// <summary>
    /// Bottom border style
    /// </summary>
    public ExcelDxfBorderItem Bottom
    {
        get;
        internal set;
    }
    /// <summary>
    /// Horizontal border style
    /// </summary>
    public ExcelDxfBorderItem Horizontal
    {
        get;
        internal set;
    }
    /// <summary>
    /// Vertical border style
    /// </summary>
    public ExcelDxfBorderItem Vertical
    {
        get;
        internal set;
    }

    /// <summary>
    /// The Id
    /// </summary>
    internal override string Id
    {
        get
        {
            return this.Top.Id + this.Bottom.Id + this.Left.Id + this.Right.Id + this.Vertical.Id + this.Horizontal.Id;
        }
    }

    /// <summary>
    /// Creates the the xml node
    /// </summary>
    /// <param name="helper">The xml helper</param>
    /// <param name="path">The X Path</param>
    internal override void CreateNodes(XmlHelper helper, string path)
    {
        this.Left.CreateNodes(helper, path + "/d:left");
        this.Right.CreateNodes(helper, path + "/d:right");
        this.Top.CreateNodes(helper, path + "/d:top");
        this.Bottom.CreateNodes(helper, path + "/d:bottom");
        this.Vertical.CreateNodes(helper, path + "/d:vertical");
        this.Horizontal.CreateNodes(helper, path + "/d:horizontal");
    }
    internal override void SetStyle()
    {
        if (this._callback != null)
        {
            this.Left.SetStyle();
            this.Right.SetStyle();
            this.Top.SetStyle();
            this.Bottom.SetStyle();
        }
    }

    /// <summary>
    /// If the object has any properties set
    /// </summary>
    public override bool HasValue
    {
        get 
        {
            return this.Left.HasValue || this.Right.HasValue || this.Top.HasValue || this.Bottom.HasValue|| this.Vertical.HasValue || this.Horizontal.HasValue;
        }
    }
    /// <summary>
    /// Clears all properties
    /// </summary>
    public override void Clear()
    {
        this.Left.Clear();
        this.Right.Clear();
        this.Top.Clear();
        this.Bottom.Clear();
        this.Vertical.Clear();
        this.Horizontal.Clear();
    }

    /// <summary>
    /// Set the border properties for Top/Bottom/Right and Left.
    /// </summary>
    /// <param name="borderStyle">The border style</param>
    /// <param name="themeColor">The theme color</param>
    public void BorderAround(ExcelBorderStyle borderStyle = ExcelBorderStyle.Thin, eThemeSchemeColor themeColor=eThemeSchemeColor.Accent1)
    {
        this.Top.Style = borderStyle;
        this.Top.Color.SetColor(themeColor);
        this.Right.Style = borderStyle;
        this.Right.Color.SetColor(themeColor);
        this.Bottom.Style = borderStyle;
        this.Bottom.Color.SetColor(themeColor);
        this.Left.Style = borderStyle;
        this.Left.Color.SetColor(themeColor);
    }
    /// <summary>
    /// Set the border properties for Top/Bottom/Right and Left.
    /// </summary>
    /// <param name="borderStyle">The border style</param>
    /// <param name="color">The color to use</param>
    public void BorderAround(ExcelBorderStyle borderStyle, Color color)
    {
        this.Top.Style = borderStyle;
        this.Top.Color.SetColor(color);
        this.Right.Style = borderStyle;
        this.Right.Color.SetColor(color);
        this.Bottom.Style = borderStyle;
        this.Bottom.Color.SetColor(color);
        this.Left.Style = borderStyle;
        this.Left.Color.SetColor(color);
    }

    /// <summary>
    /// Clone the object
    /// </summary>
    /// <returns>A new instance of the object</returns>
    internal override DxfStyleBase Clone()
    {
        return new ExcelDxfBorderBase(this._styles, this._callback) 
        { 
            Bottom = (ExcelDxfBorderItem)this.Bottom.Clone(), 
            Top= (ExcelDxfBorderItem)this.Top.Clone(), 
            Left= (ExcelDxfBorderItem)this.Left.Clone(), 
            Right= (ExcelDxfBorderItem)this.Right.Clone(),
            Vertical = (ExcelDxfBorderItem)this.Vertical.Clone(),
            Horizontal = (ExcelDxfBorderItem)this.Horizontal.Clone(),
        };
    }
    internal override void SetValuesFromXml(XmlHelper helper)
    {
        if (helper.ExistsNode("d:border"))
        {
            this.Left = this.GetBorderItem(helper, "d:border/d:left", eStyleClass.BorderLeft);
            this.Right = this.GetBorderItem(helper, "d:border/d:right", eStyleClass.BorderLeft);
            this.Bottom = this.GetBorderItem(helper, "d:border/d:bottom", eStyleClass.BorderLeft);
            this.Top = this.GetBorderItem(helper, "d:border/d:top", eStyleClass.BorderLeft);
            this.Vertical = this.GetBorderItem(helper, "d:border/d:vertical", eStyleClass.Border);
            this.Horizontal = this.GetBorderItem(helper, "d:border/d:horizontal", eStyleClass.Border);
        }
    }
    private ExcelDxfBorderItem GetBorderItem(XmlHelper helper, string path, eStyleClass styleClass)
    {
        ExcelDxfBorderItem bi = new ExcelDxfBorderItem(this._styles, styleClass, this._callback);
        bool exists = helper.ExistsNode(path);
        if (exists)
        {
            string? style = helper.GetXmlNodeString(path + "/@style");
            bi.Style = GetBorderStyleEnum(style);
            bi.Color = this.GetColor(helper, path + "/d:color", styleClass);
        }
        return bi;
    }
    private static ExcelBorderStyle? GetBorderStyleEnum(string style)
    {
        if (style == "")
        {
            return null;
        }

        string sInStyle = style.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + style.Substring(1, style.Length - 1);
        try
        {
            return (ExcelBorderStyle)Enum.Parse(typeof(ExcelBorderStyle), sInStyle);
        }
        catch
        {
            return ExcelBorderStyle.None;
        }

    }

}