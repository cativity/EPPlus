using System;
using System.Xml;
using OfficeOpenXml.Utils.Extensions;
namespace OfficeOpenXml.Drawing.Vml;

/// <summary>
/// Fill settings for a vml gradient fill
/// </summary>
public class ExcelVmlDrawingGradientFill : XmlHelper
{
    ExcelVmlDrawingFill _fill;
    internal ExcelVmlDrawingGradientFill(ExcelVmlDrawingFill fill, XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
    {
        this._fill = fill;
    }       
    /// <summary>
    /// A semicolon separated list of colors used for gradient fill. 
    /// Each color item starts with a percent and a color. Starting from 0% and ending and 100%.
    /// Use <seealso cref="SetGradientColors(VmlGradiantColor[])"/>  to set this property.
    /// </summary>
    public string ColorsString
    {
        get
        {
            return this.GetXmlNodeString("v:fill/@colors");
        }
        set
        {
            this.SetXmlNodeString("v:fill/@colors", value);
        }
    }
    /// <summary>
    /// Sets the <see cref="ColorsString"/> with the colors supplied and optionally 
    /// <see cref="ExcelVmlDrawingFill.Color"/> and <see cref="ExcelVmlDrawingFill.SecondColor"/>.
    /// Each color item starts with a percent and a color. 
    /// Percent values must be sorted, starting from 0% and ending and 100%.
    /// If 0% is omitted, <see cref="ExcelVmlDrawingFill.Color"/> is used.
    /// If 100% is omitted, <see cref="ExcelVmlDrawingFill.SecondColor"/> is used.
    /// </summary>
    /// <param name="colors">The colors with a percent value for the gradient fill</param>
    public void SetGradientColors(params VmlGradiantColor[] colors)
    {
        if(colors==null || colors.Length==0)
        {
            throw (new ArgumentException("Please supply a list of colors"));
        }
        double p = -1;
        foreach(VmlGradiantColor? c in colors)
        {
            if(c.Percent<=p)
            {
                throw (new ArgumentException("Percent values in the color list must be sorted and must be unique."));
            }
            p = c.Percent;
        }

        string? colorsString = "";
        if(colors[0].Percent!=0)
        {
            colorsString = $"0 #{colors[0].Color.ToColorString()};";
        }

        foreach(VmlGradiantColor? c in colors)
        {
            double v = c.Percent == 0 ? 0 : c.Percent / 100;
            colorsString += $"{(v * 0x10000):F0}f #{c.Color.ToColorString()};";
        }
        if(colors[0].Percent==0)
        {
            this._fill.Color.SetColor(colors[0].Color);
        }
        else if(!string.IsNullOrEmpty(this._fill.Color.ColorString))
        {
            colorsString = $"0 #{this._fill.Color.ColorString};";
        }
            
        if(colors[colors.Length-1].Percent==100)
        {
            this._fill.SecondColor.SetColor(colors[colors.Length - 1].Color);
        }
        else if (!string.IsNullOrEmpty(this._fill.SecondColor.ColorString))
        {
            colorsString += $"1 #{this._fill.SecondColor.ColorString};";
        }

        this.ColorsString = colorsString;
    }
    /// <summary>
    /// Gradient angle
    /// </summary>
    public double? Angle
    {
        get
        {
            return this.GetXmlNodeDoubleNull("v:fill/@angle");
        }
        set
        {
            this.SetXmlNodeDouble("v:fill/@angle", value);
        }
    }
    /// <summary>
    /// Gradient center
    /// </summary>
    public double? Focus
    {
        get
        {
            return this.GetXmlNodeDoubleNull("v:fill/@focus");
        }
        set
        {
            this.SetXmlNodeDouble("v:fill/@focus", value);
        }
    }
    /// <summary>
    /// Gradient method
    /// </summary>
    public eVmlGradientMethod Method
    {
        get
        {
            return this.GetXmlNodeString("v:fill/@method").ToGradientMethodEnum(eVmlGradientMethod.None);
        }
        set
        {
            this.SetXmlNodeString("v:fill/@focus", value.ToEnumString());
        }
    }

}