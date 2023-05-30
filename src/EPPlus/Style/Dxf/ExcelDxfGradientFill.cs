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

using OfficeOpenXml.Core;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf;

/// <summary>
/// Represents a gradient fill used for differential style formatting.
/// </summary>
public class ExcelDxfGradientFill : DxfStyleBase
{
    internal ExcelDxfGradientFill(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
        : base(styles, callback) =>
        this.Colors = new ExcelDxfGradientFillColorCollection(styles, callback);

    /// <summary>
    /// If the object has any properties set
    /// </summary>
    public override bool HasValue =>
        this.Colors.HasValue
        || this.Degree.HasValue
        || this.Left.HasValue
        || this.Right.HasValue
        || this.Top.HasValue
        || this.Bottom.HasValue
        || this.GradientType.HasValue;

    internal override string Id =>
        this.Colors.Id
        + "|"
        + GetAsString(this.Degree)
        + "|"
        + GetAsString(this.Left)
        + "|"
        + GetAsString(this.Right)
        + "|"
        + GetAsString(this.Top)
        + "|"
        + GetAsString(this.Bottom)
        + "|"
        + GetAsString(this.GradientType);

    /// <summary>
    /// Clears all properties
    /// </summary>
    public override void Clear()
    {
        this.Degree = null;
        this.Left = null;
        this.Right = null;
        this.Top = null;
        this.Bottom = null;
        this.Colors.Clear();
    }

    /// <summary>
    /// A collection of colors and percents for the gradient fill
    /// </summary>
    public ExcelDxfGradientFillColorCollection Colors { get; private set; }

    internal override DxfStyleBase Clone() =>
        new ExcelDxfGradientFill(this._styles, this._callback)
        {
            Colors = (ExcelDxfGradientFillColorCollection)this.Colors.Clone(),
            Degree = this.Degree,
            Left = this.Left,
            Right = this.Right,
            Top = this.Top,
            Bottom = this.Bottom
        };

    eDxfGradientFillType? _gradientType;

    /// <summary>
    /// Type of gradient fill
    /// </summary>
    public eDxfGradientFillType? GradientType
    {
        get => this._gradientType;
        set
        {
            this._gradientType = value;
            this._callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientType, value);
        }
    }

    double? _degree;

    /// <summary>
    /// Angle of the linear gradient
    /// </summary>
    public double? Degree
    {
        get => this._degree;
        set
        {
            this._degree = value;
            this._callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientDegree, value);
        }
    }

    double? _left;

    /// <summary>
    /// The left position of the inner rectangle (color 1). 
    /// </summary>
    public double? Left
    {
        get => this._left;
        set
        {
            this._left = value;
            this._callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientLeft, value);
        }
    }

    double? _right;

    /// <summary>
    /// The right position of the inner rectangle (color 1). 
    /// </summary>
    public double? Right
    {
        get => this._right;
        set
        {
            this._right = value;
            this._callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientRight, value);
        }
    }

    double? _top;

    /// <summary>
    /// The top position of the inner rectangle (color 1). 
    /// </summary>
    public double? Top
    {
        get => this._top;
        set
        {
            this._top = value;
            this._callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientTop, value);
        }
    }

    double? _bottom;

    /// <summary>
    /// The bottom position of the inner rectangle (color 1). 
    /// </summary>
    public double? Bottom
    {
        get => this._bottom;
        set
        {
            this._bottom = value;
            this._callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientBottom, value);
        }
    }

    internal override void CreateNodes(XmlHelper helper, string path)
    {
        XmlNode? gradNode = helper.CreateNode(path + "/d:gradientFill");
        XmlHelper? gradHelper = XmlHelperFactory.Create(helper.NameSpaceManager, gradNode);
        SetValueEnum(gradHelper, "@type", this.GradientType);
        SetValue(gradHelper, "@degree", this.Degree);
        SetValue(gradHelper, "@left", this.Left);
        SetValue(gradHelper, "@right", this.Right);
        SetValue(gradHelper, "@top", this.Top);
        SetValue(gradHelper, "@bottom", this.Bottom);

        foreach (ExcelDxfGradientFillColor? c in this.Colors)
        {
            c.CreateNodes(gradHelper, "");
        }
    }

    internal override void SetStyle()
    {
        if (this._callback != null)
        {
            this._callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientType, this._gradientType);
            this._callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientDegree, this._degree);
            this._callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientTop, this._top);
            this._callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientBottom, this._bottom);
            this._callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientLeft, this._left);
            this._callback?.Invoke(eStyleClass.GradientFill, eStyleProperty.GradientRight, this._right);

            foreach (ExcelDxfGradientFillColor? c in this.Colors)
            {
                c.SetStyle();
            }
        }
    }

    internal override void SetValuesFromXml(XmlHelper helper)
    {
        this.GradientType = helper.GetXmlNodeString("d:fill/d:gradientFill/@type").ToEnum<eDxfGradientFillType>();
        this.Degree = helper.GetXmlNodeDoubleNull("d:fill/d:gradientFill/@degree");
        this.Left = helper.GetXmlNodeDoubleNull("d:fill/d:gradientFill/@left");
        this.Right = helper.GetXmlNodeDoubleNull("d:fill/d:gradientFill/@right");
        this.Top = helper.GetXmlNodeDoubleNull("d:fill/d:gradientFill/@top");
        this.Bottom = helper.GetXmlNodeDoubleNull("d:fill/d:gradientFill/@bottom");

        foreach (XmlNode node in helper.GetNodes("d:fill/d:gradientFill/d:stop"))
        {
            XmlHelper? stopHelper = XmlHelperFactory.Create(this._styles.NameSpaceManager, node);
            ExcelDxfGradientFillColor? c = this.Colors.Add(stopHelper.GetXmlNodeDouble("@position") * 100);
            c.Color = this.GetColor(stopHelper, "d:color", c.Position == 0 ? eStyleClass.FillGradientColor1 : eStyleClass.FillGradientColor2);
        }
    }
}