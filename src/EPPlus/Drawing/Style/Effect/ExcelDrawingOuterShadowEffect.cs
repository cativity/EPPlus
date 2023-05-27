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
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Effect;

/// <summary>
/// The outer shadow effect. A shadow is applied outside the edges of the drawing.
/// </summary>
public class ExcelDrawingOuterShadowEffect : ExcelDrawingInnerShadowEffect
{
    private readonly string _shadowAlignmentPath = "{0}/@algn";
    private readonly string _rotateWithShapePath = "{0}/@rotWithShape";
    private readonly string _verticalSkewAnglePath = "{0}/@ky";
    private readonly string _horizontalSkewAnglePath = "{0}/@kx";
    private readonly string _verticalScalingFactorPath = "{0}/@sy";
    private readonly string _horizontalScalingFactorPath = "{0}/@sx";

    internal ExcelDrawingOuterShadowEffect(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, string path)
        : base(nameSpaceManager, topNode, schemaNodeOrder, path)
    {
        this._shadowAlignmentPath = string.Format(this._shadowAlignmentPath, path);
        this._rotateWithShapePath = string.Format(this._rotateWithShapePath, path);
        this._verticalSkewAnglePath = string.Format(this._verticalSkewAnglePath, path);
        this._horizontalSkewAnglePath = string.Format(this._horizontalSkewAnglePath, path);
        this._verticalScalingFactorPath = string.Format(this._verticalScalingFactorPath, path);
        this._horizontalScalingFactorPath = string.Format(this._horizontalScalingFactorPath, path);
    }

    /// <summary>
    /// The shadow alignment
    /// </summary>
    public eRectangleAlignment Alignment
    {
        get { return this.GetXmlNodeString(this._shadowAlignmentPath).TranslateRectangleAlignment(); }
        set
        {
            if (value == eRectangleAlignment.Bottom)
            {
                this.DeleteNode(this._shadowAlignmentPath);
            }
            else
            {
                this.SetXmlNodeString(this._shadowAlignmentPath, value.TranslateString());
            }
        }
    }

    /// <summary>
    /// If the shadow rotates with the shape
    /// </summary>
    public bool RotateWithShape
    {
        get { return this.GetXmlNodeBool(this._rotateWithShapePath, true); }
        set { this.SetXmlNodeBool(this._rotateWithShapePath, value, true); }
    }

    /// <summary>
    /// Horizontal skew angle.
    /// Ranges from -90 to 90 degrees 
    /// </summary>
    public double HorizontalSkewAngle
    {
        get { return this.GetXmlNodeAngel(this._horizontalSkewAnglePath); }
        set { this.SetXmlNodeAngel(this._horizontalSkewAnglePath, value, "HorizontalSkewAngle", -90, 90); }
    }

    /// <summary>
    /// Vertical skew angle.
    /// Ranges from -90 to 90 degrees 
    /// </summary>
    public double VerticalSkewAngle
    {
        get { return this.GetXmlNodeAngel(this._verticalSkewAnglePath); }
        set { this.SetXmlNodeAngel(this._verticalSkewAnglePath, value, "HorizontalSkewAngle", -90, 90); }
    }

    /// <summary>
    /// Horizontal scaling factor in percentage.
    /// A negative value causes a flip.
    /// </summary>
    public double HorizontalScalingFactor
    {
        get { return this.GetXmlNodePercentage(this._horizontalScalingFactorPath) ?? 100; }
        set { this.SetXmlNodePercentage(this._horizontalScalingFactorPath, value, true, 10000); }
    }

    /// <summary>
    /// Vertical scaling factor in percentage.
    /// A negative value causes a flip.
    /// </summary>
    public double VerticalScalingFactor
    {
        get { return this.GetXmlNodePercentage(this._verticalScalingFactorPath) ?? 100; }
        set { this.SetXmlNodePercentage(this._verticalScalingFactorPath, value, true, 10000); }
    }
}