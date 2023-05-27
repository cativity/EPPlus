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
using System.Globalization;

namespace OfficeOpenXml.Style.Dxf;

/// <summary>
/// A fill in a differential formatting record
/// </summary>
public class ExcelDxfFill : DxfStyleBase
{
    internal ExcelDxfFill(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback)
        : base(styles, callback)
    {
        this.PatternColor = new ExcelDxfColor(styles, eStyleClass.FillPatternColor, callback);
        this.BackgroundColor = new ExcelDxfColor(styles, eStyleClass.FillBackgroundColor, callback);
        this.Gradient = null;
    }

    ExcelFillStyle? _patternType;

    /// <summary>
    /// The pattern tyle
    /// </summary>
    public ExcelFillStyle? PatternType
    {
        get { return this._patternType; }
        set
        {
            if (this.Style == eDxfFillStyle.GradientFill)
            {
                throw new InvalidOperationException("Cant set Pattern Type when Style is set to GradientFill");
            }

            this._patternType = value;
            this._callback?.Invoke(eStyleClass.Fill, eStyleProperty.PatternType, value);
        }
    }

    /// <summary>
    /// The color of the pattern
    /// </summary>
    public ExcelDxfColor PatternColor { get; internal set; }

    /// <summary>
    /// The background color
    /// </summary>
    public ExcelDxfColor BackgroundColor { get; internal set; }

    /// <summary>
    /// The Id
    /// </summary>
    internal override string Id
    {
        get
        {
            if (this.Style == eDxfFillStyle.PatternFill)
            {
                return GetAsString(this.PatternType)
                       + "|"
                       + (this.PatternColor == null ? "" : this.PatternColor.Id)
                       + "|"
                       + (this.BackgroundColor == null ? "" : this.BackgroundColor.Id);
            }
            else
            {
                return this.Gradient.Id;
            }
        }
    }

    /// <summary>
    /// Fill style for a differential style record
    /// </summary>
    public eDxfFillStyle Style
    {
        get { return this.Gradient == null ? eDxfFillStyle.PatternFill : eDxfFillStyle.GradientFill; }
        set
        {
            if (value == eDxfFillStyle.PatternFill && this.Gradient != null)
            {
                this.PatternColor = new ExcelDxfColor(this._styles, eStyleClass.FillPatternColor, this._callback);
                this.BackgroundColor = new ExcelDxfColor(this._styles, eStyleClass.FillBackgroundColor, this._callback);
                this.Gradient = null;
            }
            else if (value == eDxfFillStyle.GradientFill && this.Gradient == null)
            {
                this.PatternType = null;
                this.PatternColor = null;
                this.BackgroundColor = null;
                this.Gradient = new ExcelDxfGradientFill(this._styles, this._callback);
            }
        }
    }

    /// <summary>
    /// Gradient fill settings
    /// </summary>
    public ExcelDxfGradientFill Gradient { get; internal set; }

    /// <summary>
    /// Creates the the xml node
    /// </summary>
    /// <param name="helper">The xml helper</param>
    /// <param name="path">The X Path</param>
    internal override void CreateNodes(XmlHelper helper, string path)
    {
        _ = helper.CreateNode(path);

        if (this.Style == eDxfFillStyle.PatternFill)
        {
            SetValueEnum(helper, path + "/d:patternFill/@patternType", this.PatternType);
            SetValueColor(helper, path + "/d:patternFill/d:fgColor", this.PatternColor);
            SetValueColor(helper, path + "/d:patternFill/d:bgColor", this.BackgroundColor);
        }
        else
        {
            this.Gradient.CreateNodes(helper, path);
        }
    }

    /// <summary>
    /// If the object has any properties set
    /// </summary>
    public override bool HasValue
    {
        get
        {
            if (this.Style == eDxfFillStyle.PatternFill)
            {
                return this.PatternType != null || this.PatternColor.HasValue || this.BackgroundColor.HasValue;
            }
            else
            {
                return this.Gradient.HasValue;
            }
        }
    }

    /// <summary>
    /// Clears the fill
    /// </summary>
    public override void Clear()
    {
        if (this.Style == eDxfFillStyle.PatternFill)
        {
            this.PatternType = null;
            this.PatternColor.Clear();
            this.BackgroundColor.Clear();
        }
        else
        {
            this.Gradient.Clear();
        }
    }

    /// <summary>
    /// Clone the object
    /// </summary>
    /// <returns>A new instance of the object</returns>
    internal override DxfStyleBase Clone()
    {
        return new ExcelDxfFill(this._styles, this._callback)
        {
            PatternType = this.PatternType,
            PatternColor = (ExcelDxfColor)this.PatternColor?.Clone(),
            BackgroundColor = (ExcelDxfColor)this.BackgroundColor?.Clone(),
            Gradient = (ExcelDxfGradientFill)this.Gradient?.Clone()
        };
    }

    internal override void SetStyle()
    {
        if (this._callback != null)
        {
            if (this.Style == eDxfFillStyle.PatternFill)
            {
                this._callback?.Invoke(eStyleClass.Fill, eStyleProperty.PatternType, this._patternType);
                this.PatternColor.SetStyle();
                this.BackgroundColor.SetStyle();
            }
            else
            {
                this.Gradient.SetStyle();
            }
        }
    }

    internal override void SetValuesFromXml(XmlHelper helper)
    {
        if (helper.ExistsNode("d:fill/d:patternFill"))
        {
            this.PatternType = GetPatternTypeEnum(helper.GetXmlNodeString("d:fill/d:patternFill/@patternType"));
            this.BackgroundColor = this.GetColor(helper, "d:fill/d:patternFill/d:bgColor/", eStyleClass.FillBackgroundColor);
            this.PatternColor = this.GetColor(helper, "d:fill/d:patternFill/d:fgColor/", eStyleClass.FillPatternColor);
            this.Gradient = null;
        }
        else if (helper.ExistsNode("d:fill/d:gradientFill"))
        {
            this.PatternType = null;
            this.BackgroundColor = null;
            this.PatternColor = null;
            this.Gradient = new ExcelDxfGradientFill(this._styles, this._callback);
            this.Gradient.SetValuesFromXml(helper);
        }
    }

    internal static ExcelFillStyle? GetPatternTypeEnum(string patternType)
    {
        if (string.IsNullOrEmpty(patternType))
        {
            return null;
        }

        patternType = patternType.Substring(0, 1).ToUpper(CultureInfo.InvariantCulture) + patternType.Substring(1, patternType.Length - 1);

        try
        {
            return (ExcelFillStyle)Enum.Parse(typeof(ExcelFillStyle), patternType);
        }
        catch
        {
            return null;
        }
    }
}