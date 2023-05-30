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
using OfficeOpenXml.Drawing.Interfaces;
using System.Xml;
using System.Globalization;
using System.Drawing;
using System.IO;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml.Drawing.Vml;

/// <summary>
/// Drawing object used for header and footer pictures
/// </summary>
public class ExcelVmlDrawingPicture : ExcelVmlDrawingBase, IPictureContainer
{
    ExcelWorksheet _worksheet;

    internal ExcelVmlDrawingPicture(XmlNode topNode, XmlNamespaceManager ns, ExcelWorksheet ws)
        : base(topNode, ns) =>
        this._worksheet = ws;

    /// <summary>
    /// Position ID
    /// </summary>
    public string Position => this.GetXmlNodeString("@id");

    /// <summary>
    /// The width in points
    /// </summary>
    public double Width
    {
        get => this.GetStyleProp("width");
        set => this.SetStyleProp("width", value.ToString(CultureInfo.InvariantCulture) + "pt");
    }

    /// <summary>
    /// The height in points
    /// </summary>
    public double Height
    {
        get => this.GetStyleProp("height");
        set => this.SetStyleProp("height", value.ToString(CultureInfo.InvariantCulture) + "pt");
    }

    /// <summary>
    /// Margin Left in points
    /// </summary>
    public double Left
    {
        get => this.GetStyleProp("left");
        set => this.SetStyleProp("left", value.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// Margin top in points
    /// </summary>
    public double Top
    {
        get => this.GetStyleProp("top");
        set => this.SetStyleProp("top", value.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// The Title of the image
    /// </summary>
    public string Title
    {
        get => this.GetXmlNodeString("v:imagedata/@o:title");
        set => this.SetXmlNodeString("v:imagedata/@o:title", value);
    }

    ExcelImage _image;

    /// <summary>
    /// The Image
    /// </summary>
    public ExcelImage Image
    {
        get
        {
            if (this._image == null)
            {
                this._image = new ExcelImage(this, new ePictureType[] { ePictureType.Svg, ePictureType.Ico, ePictureType.WebP });
                ZipPackage? pck = this._worksheet._package.ZipPackage;

                if (pck.PartExists(this.ImageUri))
                {
                    ZipPackagePart? part = pck.GetPart(this.ImageUri);
                    _ = this._image.SetImage(((MemoryStream)part.GetStream()).ToArray(), PictureStore.GetPictureType(this.ImageUri));
                }
                else
                {
                    return null;
                }
            }

            return this._image;
        }
    }

    internal Uri ImageUri { get; set; }

    internal string RelId
    {
        get => this.GetXmlNodeString("v:imagedata/@o:relid");
        set => this.SetXmlNodeString("v:imagedata/@o:relid", value);
    }

    /// <summary>
    /// Determines whether an image will be displayed in black and white
    /// </summary>
    public bool BiLevel
    {
        get => this.GetXmlNodeString("v:imagedata/@bilevel") == "t";
        set
        {
            if (value)
            {
                this.SetXmlNodeString("v:imagedata/@bilevel", "t");
            }
            else
            {
                this.DeleteNode("v:imagedata/@bilevel");
            }
        }
    }

    /// <summary>
    /// Determines whether a picture will be displayed in grayscale mode
    /// </summary>
    public bool GrayScale
    {
        get => this.GetXmlNodeString("v:imagedata/@grayscale") == "t";
        set
        {
            if (value)
            {
                this.SetXmlNodeString("v:imagedata/@grayscale", "t");
            }
            else
            {
                this.DeleteNode("v:imagedata/@grayscale");
            }
        }
    }

    /// <summary>
    /// Defines the intensity of all colors in an image
    /// Default value is 1
    /// </summary>
    public double Gain
    {
        get
        {
            string v = this.GetXmlNodeString("v:imagedata/@gain");

            return GetFracDT(v, 1);
        }
        set
        {
            if (value < 0)
            {
                throw new ArgumentOutOfRangeException("Value must be positive");
            }

            if (value == 1)
            {
                this.DeleteNode("v:imagedata/@gamma");
            }
            else
            {
                this.SetXmlNodeString("v:imagedata/@gain", value.ToString("#.0#", CultureInfo.InvariantCulture));
            }
        }
    }

    /// <summary>
    /// Defines the amount of contrast for an image
    /// Default value is 0;
    /// </summary>
    public double Gamma
    {
        get
        {
            string v = this.GetXmlNodeString("v:imagedata/@gamma");

            return GetFracDT(v, 0);
        }
        set
        {
            if (value == 0) //Default
            {
                this.DeleteNode("v:imagedata/@gamma");
            }
            else
            {
                this.SetXmlNodeString("v:imagedata/@gamma", value.ToString("#.0#", CultureInfo.InvariantCulture));
            }
        }
    }

    /// <summary>
    /// Defines the intensity of black in an image
    /// Default value is 0
    /// </summary>
    public double BlackLevel
    {
        get
        {
            string v = this.GetXmlNodeString("v:imagedata/@blacklevel");

            return GetFracDT(v, 0);
        }
        set
        {
            if (value == 0)
            {
                this.DeleteNode("v:imagedata/@blacklevel");
            }
            else
            {
                this.SetXmlNodeString("v:imagedata/@blacklevel", value.ToString("#.0#", CultureInfo.InvariantCulture));
            }
        }
    }

    #region Private Methods

    private static double GetFracDT(string v, double def)
    {
        double d;

        if (v.EndsWith("f"))
        {
            v = v.Substring(0, v.Length - 1);

            if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
            {
                d /= 65535;
            }
            else
            {
                d = def;
            }
        }
        else
        {
            if (!double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
            {
                d = def;
            }
        }

        return d;
    }

    private void SetStyleProp(string propertyName, string value)
    {
        string style = this.GetXmlNodeString("@style");
        string newStyle = "";
        bool found = false;

        foreach (string prop in style.Split(';'))
        {
            string[] split = prop.Split(':');

            if (split[0] == propertyName)
            {
                newStyle += propertyName + ":" + value + ";";
                found = true;
            }
            else
            {
                newStyle += prop + ";";
            }
        }

        if (!found)
        {
            newStyle += propertyName + ":" + value + ";";
        }

        this.SetXmlNodeString("@style", newStyle.Substring(0, newStyle.Length - 1));
    }

    private double GetStyleProp(string propertyName)
    {
        string style = this.GetXmlNodeString("@style");

        foreach (string prop in style.Split(';'))
        {
            string[] split = prop.Split(':');

            if (split[0] == propertyName && split.Length > 1)
            {
                string value = split[1].EndsWith("pt") ? split[1].Substring(0, split[1].Length - 2) : split[1];

                if (double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out double ret))
                {
                    return ret;
                }
                else
                {
                    return 0;
                }
            }
        }

        return 0;
    }

    IPictureRelationDocument RelationDocument => this._worksheet.VmlDrawings;

    string ImageHash { get; set; }

    Uri UriPic { get; set; }

    ZipPackageRelationship RelPic { get; set; }

    IPictureRelationDocument IPictureContainer.RelationDocument => throw new NotImplementedException();

    string IPictureContainer.ImageHash
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    Uri IPictureContainer.UriPic
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    ZipPackageRelationship IPictureContainer.RelPic
    {
        get => throw new NotImplementedException();
        set => throw new NotImplementedException();
    }

    void IPictureContainer.RemoveImage()
    {
    }

    void IPictureContainer.SetNewImage()
    {
    }

    #endregion
}