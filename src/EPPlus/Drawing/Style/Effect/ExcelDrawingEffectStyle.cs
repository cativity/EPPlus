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

using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Effect;

/// <summary>
/// Effect styles of a drawing object
/// </summary>
public class ExcelDrawingEffectStyle : XmlHelper
{
    private readonly string _path;
    private readonly string _softEdgeRadiusPath = "{0}a:softEdge/@rad";
    private readonly string _blurPath = "{0}a:blur";
    private readonly string _fillOverlayPath = "{0}a:fillOverlay";
    private readonly string _glowPath = "{0}a:glow";
    private readonly string _innerShadowPath = "{0}a:innerShdw";
    private readonly string _outerShadowPath = "{0}a:outerShdw";
    private readonly string _presetShadowPath = "{0}a:prstShdw";
    private readonly string _reflectionPath = "{0}a:reflection";
    private readonly IPictureRelationDocument _pictureRelationDocument;

    internal ExcelDrawingEffectStyle(IPictureRelationDocument pictureRelationDocument,
                                     XmlNamespaceManager nameSpaceManager,
                                     XmlNode topNode,
                                     string path,
                                     string[] schemaNodeOrder)
        : base(nameSpaceManager, topNode)
    {
        this._path = path;

        if (path.Length > 0 && !path.EndsWith("/"))
        {
            path += "/";
        }

        this._softEdgeRadiusPath = string.Format(this._softEdgeRadiusPath, path);
        this._blurPath = string.Format(this._blurPath, path);
        this._fillOverlayPath = string.Format(this._fillOverlayPath, path);
        this._glowPath = string.Format(this._glowPath, path);
        this._innerShadowPath = string.Format(this._innerShadowPath, path);
        this._outerShadowPath = string.Format(this._outerShadowPath, path);
        this._presetShadowPath = string.Format(this._presetShadowPath, path);
        this._reflectionPath = string.Format(this._reflectionPath, path);
        this._pictureRelationDocument = pictureRelationDocument;

        this.AddSchemaNodeOrder(schemaNodeOrder, ExcelShapeBase._shapeNodeOrder);
    }

    ExcelDrawingBlurEffect _blur = null;

    /// <summary>
    /// The blur effect
    /// </summary>
    public ExcelDrawingBlurEffect Blur
    {
        get { return this._blur ??= new ExcelDrawingBlurEffect(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, this._blurPath); }
    }

    ExcelDrawingFillOverlayEffect _fillOverlay = null;

    /// <summary>
    /// The fill overlay effect. A fill overlay can be used to specify an additional fill for a drawing and blend the two together.
    /// </summary>
    public ExcelDrawingFillOverlayEffect FillOverlay
    {
        get
        {
            return this._fillOverlay ??= new ExcelDrawingFillOverlayEffect(this._pictureRelationDocument,
                                                                           this.NameSpaceManager,
                                                                           this.TopNode,
                                                                           this.SchemaNodeOrder,
                                                                           this._fillOverlayPath);
        }
    }

    ExcelDrawingGlowEffect _glow = null;

    /// <summary>
    /// The glow effect. A color blurred outline is added outside the edges of the drawing
    /// </summary>
    public ExcelDrawingGlowEffect Glow
    {
        get { return this._glow ??= new ExcelDrawingGlowEffect(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, this._glowPath); }
    }

    ExcelDrawingInnerShadowEffect _innerShadowEffect = null;

    /// <summary>
    /// The inner shadow effect. A shadow is applied within the edges of the drawing.
    /// </summary>
    public ExcelDrawingInnerShadowEffect InnerShadow
    {
        get
        {
            return this._innerShadowEffect ??=
                       new ExcelDrawingInnerShadowEffect(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, this._innerShadowPath);
        }
    }

    ExcelDrawingOuterShadowEffect _outerShadow = null;

    /// <summary>
    /// The outer shadow effect. A shadow is applied outside the edges of the drawing.
    /// </summary>
    public ExcelDrawingOuterShadowEffect OuterShadow
    {
        get
        {
            return this._outerShadow ??= new ExcelDrawingOuterShadowEffect(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, this._outerShadowPath);
        }
    }

    ExcelDrawingPresetShadowEffect _presetShadow;

    /// <summary>
    /// The preset shadow effect.
    /// </summary>
    public ExcelDrawingPresetShadowEffect PresetShadow
    {
        get
        {
            return this._presetShadow ??= new ExcelDrawingPresetShadowEffect(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, this._presetShadowPath);
        }
    }

    ExcelDrawingReflectionEffect _reflection;

    /// <summary>
    /// The reflection effect.
    /// </summary>
    public ExcelDrawingReflectionEffect Reflection
    {
        get { return this._reflection ??= new ExcelDrawingReflectionEffect(this.NameSpaceManager, this.TopNode, this.SchemaNodeOrder, this._reflectionPath); }
    }

    /// <summary>
    /// Soft edge radius. A null value indicates no radius
    /// </summary>
    public double? SoftEdgeRadius
    {
        get { return this.GetXmlNodeEmuToPtNull(this._softEdgeRadiusPath); }
        set
        {
            if (value.HasValue)
            {
                this.SetXmlNodeEmuToPt(this._softEdgeRadiusPath, value.Value);
            }
            else
            {
                this.DeleteNode(this._softEdgeRadiusPath, true);
            }
        }
    }

    internal XmlElement EffectElement
    {
        get
        {
            if (string.IsNullOrEmpty(this._path))
            {
                return (XmlElement)this.TopNode;
            }

            if (this.ExistsNode(this._path))
            {
                return (XmlElement)this.GetNode(this._path);
            }

            return (XmlElement)this.CreateNode(this._path);
        }
    }

    /// <summary>
    /// If the drawing has any inner shadow properties set
    /// </summary>
    public bool HasInnerShadow
    {
        get { return this.ExistsNode(this._innerShadowPath); }
    }

    /// <summary>
    /// If the drawing has any outer shadow properties set
    /// </summary>
    public bool HasOuterShadow
    {
        get { return this.ExistsNode(this._outerShadowPath); }
    }

    /// <summary>
    /// If the drawing has any preset shadow properties set
    /// </summary>
    public bool HasPresetShadow
    {
        get { return this.ExistsNode(this._presetShadowPath); }
    }

    /// <summary>
    /// If the drawing has any blur properties set
    /// </summary>
    public bool HasBlur
    {
        get { return this.ExistsNode(this._blurPath); }
    }

    /// <summary>
    /// If the drawing has any glow properties set
    /// </summary>
    public bool HasGlow
    {
        get { return this.ExistsNode(this._glowPath); }
    }

    /// <summary>
    /// If the drawing has any fill overlay properties set
    /// </summary>
    public bool HasFillOverlay
    {
        get { return this.ExistsNode(this._fillOverlayPath); }
    }

    internal void SetFromXml(XmlElement copyFromEffectElement)
    {
        XmlElement effectElement = this.EffectElement;

        foreach (XmlAttribute a in copyFromEffectElement.Attributes)
        {
            effectElement.SetAttribute(a.Name, a.NamespaceURI, a.Value);
        }

        effectElement.InnerXml = copyFromEffectElement.InnerXml;
    }

    #region Private Methods

    private void SetPredefinedOuterShadow(ePresetExcelShadowType shadowType)
    {
        this.OuterShadow.Color.SetPresetColor(ePresetColor.Black);

        switch (shadowType)
        {
            case ePresetExcelShadowType.PerspectiveUpperLeft:
                this.OuterShadow.Color.Transforms.AddAlpha(20);
                this.OuterShadow.BlurRadius = 6;
                this.OuterShadow.Distance = 0;
                this.OuterShadow.Direction = 225;
                this.OuterShadow.Alignment = eRectangleAlignment.BottomRight;
                this.OuterShadow.HorizontalSkewAngle = 20;
                this.OuterShadow.VerticalScalingFactor = 23;

                break;

            case ePresetExcelShadowType.PerspectiveUpperRight:
                this.OuterShadow.Color.Transforms.AddAlpha(20);
                this.OuterShadow.BlurRadius = 6;
                this.OuterShadow.Distance = 0;
                this.OuterShadow.Direction = 315;
                this.OuterShadow.Alignment = eRectangleAlignment.BottomLeft;
                this.OuterShadow.HorizontalSkewAngle = -20;
                this.OuterShadow.VerticalScalingFactor = 23;

                break;

            case ePresetExcelShadowType.PerspectiveBelow:
                this.OuterShadow.Color.Transforms.AddAlpha(15);
                this.OuterShadow.BlurRadius = 12;
                this.OuterShadow.Distance = 25;
                this.OuterShadow.Direction = 90;
                this.OuterShadow.HorizontalScalingFactor = 90;
                this.OuterShadow.VerticalScalingFactor = -19;

                break;

            case ePresetExcelShadowType.PerspectiveLowerLeft:
                this.OuterShadow.Color.Transforms.AddAlpha(20);
                this.OuterShadow.BlurRadius = 6;
                this.OuterShadow.Distance = 1;
                this.OuterShadow.Direction = 135;
                this.OuterShadow.Alignment = eRectangleAlignment.BottomRight;
                this.OuterShadow.HorizontalSkewAngle = 13.34;
                this.OuterShadow.VerticalScalingFactor = -23;

                break;

            case ePresetExcelShadowType.PerspectiveLowerRight:
                this.OuterShadow.Color.Transforms.AddAlpha(20);
                this.OuterShadow.BlurRadius = 6;
                this.OuterShadow.Distance = 1;
                this.OuterShadow.Direction = 45;
                this.OuterShadow.Alignment = eRectangleAlignment.BottomLeft;
                this.OuterShadow.HorizontalSkewAngle = -13.34;
                this.OuterShadow.VerticalScalingFactor = -23;

                break;

            case ePresetExcelShadowType.OuterCenter:
                this.OuterShadow.Color.Transforms.AddAlpha(40);
                this.OuterShadow.VerticalScalingFactor = 102;
                this.OuterShadow.HorizontalScalingFactor = 102;
                this.OuterShadow.BlurRadius = 5;
                this.OuterShadow.Alignment = eRectangleAlignment.Center;

                break;

            default:
                this.OuterShadow.Color.Transforms.AddAlpha(40);
                this.OuterShadow.BlurRadius = 4;
                this.OuterShadow.Distance = 3;

                switch (shadowType)
                {
                    case ePresetExcelShadowType.OuterTopLeft:
                        this.OuterShadow.Direction = 225;
                        this.OuterShadow.Alignment = eRectangleAlignment.BottomRight;

                        break;

                    case ePresetExcelShadowType.OuterTop:
                        this.OuterShadow.Direction = 270;
                        this.OuterShadow.Alignment = eRectangleAlignment.Bottom;

                        break;

                    case ePresetExcelShadowType.OuterTopRight:
                        this.OuterShadow.Direction = 315;
                        this.OuterShadow.Alignment = eRectangleAlignment.BottomLeft;

                        break;

                    case ePresetExcelShadowType.OuterLeft:
                        this.OuterShadow.Direction = 180;
                        this.OuterShadow.Alignment = eRectangleAlignment.Right;

                        break;

                    case ePresetExcelShadowType.OuterRight:
                        this.OuterShadow.Direction = 0;
                        this.OuterShadow.Alignment = eRectangleAlignment.Left;

                        break;

                    case ePresetExcelShadowType.OuterBottomLeft:
                        this.OuterShadow.Direction = 135;
                        this.OuterShadow.Alignment = eRectangleAlignment.TopRight;

                        break;

                    case ePresetExcelShadowType.OuterBottom:
                        this.OuterShadow.Direction = 90;
                        this.OuterShadow.Alignment = eRectangleAlignment.Top;

                        break;

                    case ePresetExcelShadowType.OuterBottomRight:
                        this.OuterShadow.Direction = 45;
                        this.OuterShadow.Alignment = eRectangleAlignment.TopLeft;

                        break;
                }

                break;
        }

        this.OuterShadow.RotateWithShape = false;
    }

    private void SetPredefinedInnerShadow(ePresetExcelShadowType shadowType)
    {
        this.InnerShadow.Color.SetPresetColor(ePresetColor.Black);

        if (shadowType == ePresetExcelShadowType.InnerCenter)
        {
            this.InnerShadow.Color.Transforms.AddAlpha(0);
            this.InnerShadow.Direction = 0;
            this.InnerShadow.Distance = 0;
            this.InnerShadow.BlurRadius = 9;
        }
        else
        {
            this.InnerShadow.Color.Transforms.AddAlpha(50);
            this.InnerShadow.BlurRadius = 5;
            this.InnerShadow.Distance = 4;
        }

        switch (shadowType)
        {
            case ePresetExcelShadowType.InnerTopLeft:
                this.InnerShadow.Direction = 225;

                break;

            case ePresetExcelShadowType.InnerTop:
                this.InnerShadow.Direction = 270;

                break;

            case ePresetExcelShadowType.InnerTopRight:
                this.InnerShadow.Direction = 315;

                break;

            case ePresetExcelShadowType.InnerLeft:
                this.InnerShadow.Direction = 180;

                break;

            case ePresetExcelShadowType.InnerRight:
                this.InnerShadow.Direction = 0;

                break;

            case ePresetExcelShadowType.InnerBottomLeft:
                this.InnerShadow.Direction = 135;

                break;

            case ePresetExcelShadowType.InnerBottom:
                this.InnerShadow.Direction = 90;

                break;

            case ePresetExcelShadowType.InnerBottomRight:
                this.InnerShadow.Direction = 45;

                break;
        }
    }

    /// <summary>
    /// Set a predefined glow matching the preset types in Excel
    /// </summary>
    /// <param name="softEdgesType">The preset type</param>
    public void SetPresetSoftEdges(ePresetExcelSoftEdgesType softEdgesType)
    {
        switch (softEdgesType)
        {
            case ePresetExcelSoftEdgesType.None:
                this.SoftEdgeRadius = null;

                break;

            case ePresetExcelSoftEdgesType.SoftEdge1Pt:
                this.SoftEdgeRadius = 1;

                break;

            case ePresetExcelSoftEdgesType.SoftEdge2_5Pt:
                this.SoftEdgeRadius = 2.5;

                break;

            case ePresetExcelSoftEdgesType.SoftEdge5Pt:
                this.SoftEdgeRadius = 5;

                break;

            case ePresetExcelSoftEdgesType.SoftEdge10Pt:
                this.SoftEdgeRadius = 10;

                break;

            case ePresetExcelSoftEdgesType.SoftEdge25Pt:
                this.SoftEdgeRadius = 25;

                break;

            case ePresetExcelSoftEdgesType.SoftEdge50Pt:
                this.SoftEdgeRadius = 50;

                break;
        }
    }

    /// <summary>
    /// Set a predefined glow matching the preset types in Excel
    /// </summary>
    /// <param name="glowType">The preset type</param>
    public void SetPresetGlow(ePresetExcelGlowType glowType)
    {
        this.Glow.Delete();

        if (glowType == ePresetExcelGlowType.None)
        {
            return;
        }

        string? glowTypeString = glowType.ToString();
        string? font = glowTypeString.Substring(0, glowTypeString.IndexOf('_'));
        eSchemeColor schemeColor = (eSchemeColor)Enum.Parse(typeof(eSchemeColor), font);
        this.Glow.Color.SetSchemeColor(schemeColor);
        this.Glow.Color.Transforms.AddAlpha(40);
        this.Glow.Color.Transforms.AddSaturationModulation(175);
        this.Glow.Radius = int.Parse(glowTypeString.Substring(font.Length + 1, glowTypeString.Length - font.Length - 3));
    }

    /// <summary>
    /// Set a predefined shadow matching the preset types in Excel
    /// </summary>
    /// <param name="shadowType">The preset type</param>
    public void SetPresetShadow(ePresetExcelShadowType shadowType)
    {
        this.InnerShadow.Delete();
        this.OuterShadow.Delete();
        this.PresetShadow.Delete();

        if (shadowType == ePresetExcelShadowType.None)
        {
            return;
        }

        if (shadowType <= ePresetExcelShadowType.InnerBottomRight)
        {
            this.SetPredefinedInnerShadow(shadowType);
        }
        else
        {
            this.SetPredefinedOuterShadow(shadowType);
        }
    }

    /// <summary>
    /// Set a predefined glow matching the preset types in Excel
    /// </summary>
    /// <param name="reflectionType">The preset type</param>
    public void SetPresetReflection(ePresetExcelReflectionType reflectionType)
    {
        this.Reflection.Delete();

        if (reflectionType == ePresetExcelReflectionType.TightTouching
            || reflectionType == ePresetExcelReflectionType.Tight4Pt
            || reflectionType == ePresetExcelReflectionType.Tight8Pt)
        {
            this.Reflection.Alignment = eRectangleAlignment.BottomLeft;
            this.Reflection.RotateWithShape = false;
            this.Reflection.Direction = 90;
            this.Reflection.VerticalScalingFactor = -100;
            this.Reflection.BlurRadius = 0.5;

            if (reflectionType == ePresetExcelReflectionType.TightTouching)
            {
                this.Reflection.EndPosition = 35;
                this.Reflection.StartOpacity = 52;
                this.Reflection.EndOpacity = 0.3;
                this.Reflection.Distance = 0;
            }
            else if (reflectionType == ePresetExcelReflectionType.Tight4Pt)
            {
                this.Reflection.EndPosition = 38.5;
                this.Reflection.StartOpacity = 50;
                this.Reflection.EndOpacity = 0.3;
                this.Reflection.Distance = 4;
            }
            else if (reflectionType == ePresetExcelReflectionType.Tight8Pt)
            {
                this.Reflection.EndPosition = 40;
                this.Reflection.StartOpacity = 50;
                this.Reflection.EndOpacity = 0.275;
                this.Reflection.Distance = 8;
            }
        }
        else if (reflectionType == ePresetExcelReflectionType.HalfTouching
                 || reflectionType == ePresetExcelReflectionType.Half4Pt
                 || reflectionType == ePresetExcelReflectionType.Half8Pt)
        {
            this.Reflection.Alignment = eRectangleAlignment.BottomLeft;
            this.Reflection.RotateWithShape = false;
            this.Reflection.Direction = 90;
            this.Reflection.VerticalScalingFactor = -100;
            this.Reflection.BlurRadius = 0.5;

            if (reflectionType == ePresetExcelReflectionType.HalfTouching)
            {
                this.Reflection.EndPosition = 55;
                this.Reflection.StartOpacity = 50;
                this.Reflection.EndOpacity = 0.3;
                this.Reflection.Distance = 0;
            }
            else if (reflectionType == ePresetExcelReflectionType.Half4Pt)
            {
                this.Reflection.EndPosition = 55.5;
                this.Reflection.StartOpacity = 50;
                this.Reflection.EndOpacity = 0.300;
                this.Reflection.Distance = 4;
            }
            else if (reflectionType == ePresetExcelReflectionType.Half8Pt)
            {
                this.Reflection.EndPosition = 55.5;
                this.Reflection.StartOpacity = 50;
                this.Reflection.EndOpacity = 0.300;
                this.Reflection.Distance = 8;
            }
        }
        else if (reflectionType == ePresetExcelReflectionType.FullTouching
                 || reflectionType == ePresetExcelReflectionType.Full4Pt
                 || reflectionType == ePresetExcelReflectionType.Full8Pt)
        {
            this.Reflection.Alignment = eRectangleAlignment.BottomLeft;
            this.Reflection.RotateWithShape = false;
            this.Reflection.Direction = 90;
            this.Reflection.VerticalScalingFactor = -100;
            this.Reflection.BlurRadius = 0.5;

            if (reflectionType == ePresetExcelReflectionType.FullTouching)
            {
                this.Reflection.EndPosition = 90;
                this.Reflection.StartOpacity = 50;
                this.Reflection.EndOpacity = 0.3;
                this.Reflection.Distance = 0;
            }
            else if (reflectionType == ePresetExcelReflectionType.Full4Pt)
            {
                this.Reflection.EndPosition = 90;
                this.Reflection.StartOpacity = 50;
                this.Reflection.EndOpacity = 0.300;
                this.Reflection.Distance = 4;
            }
            else if (reflectionType == ePresetExcelReflectionType.Full8Pt)
            {
                this.Reflection.EndPosition = 92;
                this.Reflection.StartOpacity = 50;
                this.Reflection.EndOpacity = 0.295;
                this.Reflection.Distance = 8;
            }
        }
    }

    #endregion
}