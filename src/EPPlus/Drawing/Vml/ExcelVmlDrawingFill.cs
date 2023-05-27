using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml;

/// <summary>
/// Fill settings for a vml drawing
/// </summary>
public class ExcelVmlDrawingFill : XmlHelper
{
    internal ExcelDrawings _drawings;
    internal ExcelVmlDrawingFill(ExcelDrawings drawings, XmlNamespaceManager ns, XmlNode topNode, string[] schemaNodeOrder) :
        base(ns, topNode)
    {
        this.SchemaNodeOrder = schemaNodeOrder;
        this._drawings = drawings;
    }
    /// <summary>
    /// The type of fill used in the vml drawing
    /// </summary>
    public eVmlFillType Style
    {
        get
        {
            return this.GetXmlNodeString("v:fill/@type").ToEnum(eVmlFillType.NoFill);
        }
        set
        {
            if (value == eVmlFillType.NoFill)
            {
                this.SetXmlNodeString("@filled", "t");
                this.DeleteNode("v:fill");
            }
            else
            {
                this.DeleteNode("@filled");
                this.SetXmlNodeString("v:fill/@type", value.ToEnumString());
            }
        }
    }
    ExcelVmlDrawingColor _fillColor = null;
    /// <summary>
    /// The primary color used for filling the drawing.
    /// </summary>
    public ExcelVmlDrawingColor Color
    {
        get { return this._fillColor ??= new ExcelVmlDrawingColor(this.NameSpaceManager, this.TopNode, "@fillcolor"); }
    }
    /// <summary>
    /// Opacity for fill color 1. Spans 0-100%. 
    /// Transparency is is 100-Opacity
    /// </summary>
    public double Opacity
    {
        get
        {
            return VmlConvertUtil.GetOpacityFromStringVml(this.GetXmlNodeString("v:fill/@opacity"));
        }
        set
        {
            if(value < 0 || value > 100)
            {
                throw new ArgumentOutOfRangeException("Opacity ranges from 0 to 100%");
            }

            this.SetXmlNodeDouble("v:fill/@opacity", value, null, "%");
        }
    }
    ExcelVmlDrawingColor _secondColor;
    /// <summary>
    /// Fill color 2. 
    /// </summary>
    public ExcelVmlDrawingColor SecondColor
    {
        get { return this._secondColor ??= new ExcelVmlDrawingColor(this.NameSpaceManager, this.TopNode, "v:fill/@color2"); }
    }
    /// <summary>
    /// Opacity for fill color 2. Spans 0-100%
    /// Transparency is is 100-Opacity
    /// </summary>
    public double SecondColorOpacity
    {
        get
        {
            return VmlConvertUtil.GetOpacityFromStringVml(this.GetXmlNodeString("v:fill/@o:opacity2"));
        }
        set
        {
            if (value < 0 || value > 100)
            {
                throw new ArgumentOutOfRangeException("Opacity ranges from 0 to 100%");
            }

            this.SetXmlNodeDouble("v:fill/@o:opacity2", value, null, "%");
        }
    }
    ExcelVmlDrawingGradientFill _gradientSettings = null;
    /// <summary>
    /// Gradient specific settings used when <see cref="Style"/> is set to Gradient or GradientRadial.
    /// </summary>
    public ExcelVmlDrawingGradientFill GradientSettings
    {
        get { return this._gradientSettings ??= new ExcelVmlDrawingGradientFill(this, this.NameSpaceManager, this.TopNode); }
    }
    internal ExcelVmlDrawingPictureFill _patternPictureSettings = null;
    /// <summary>
    /// Image and pattern specific settings used when <see cref="Style"/> is set to Pattern, Tile or Frame.
    /// </summary>
    public ExcelVmlDrawingPictureFill PatternPictureSettings
    {
        get
        {
            return this._patternPictureSettings ??= new ExcelVmlDrawingPictureFill(this, this.NameSpaceManager, this.TopNode);
        }
    }
    /// <summary>
    /// Recolor with picture
    /// </summary>
    public bool Recolor 
    { 
        get
        {
            return this.GetXmlNodeBool("v:fill/@recolor");
        }
        set
        {
            this.SetXmlNodeBoolVml("v:fill/@recolor", value);
        }
    }
    /// <summary>
    /// Rotate fill with shape
    /// </summary>
    public bool Rotate 
    {
        get
        {
            return this.GetXmlNodeBool("v:fill/@rotate");
        }
        set
        {
            this.SetXmlNodeBoolVml("v:fill/@rotate", value);
        }
    }
}