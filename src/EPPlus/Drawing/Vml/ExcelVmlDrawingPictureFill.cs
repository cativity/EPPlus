using System;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils.Extensions;
namespace OfficeOpenXml.Drawing.Vml;

/// <summary>
/// Fill settings for a vml pattern or picture fill
/// </summary>
public class ExcelVmlDrawingPictureFill : XmlHelper, IPictureContainer
{
    ExcelVmlDrawingFill _fill;
    internal ExcelVmlDrawingPictureFill(ExcelVmlDrawingFill fill, XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
    {
        this._fill = fill;
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
    /// <summary>
    /// The aspect ratio 
    /// </summary>
    public eVmlAspectRatio AspectRatio 
    { 
        get
        {
            return this.GetXmlNodeString("v:fill/@aspect").ToEnum(eVmlAspectRatio.Ignore);
        }
        set
        {
            this.SetXmlNodeString("v:fill/@aspect", value.ToString().ToLower());
        }
    }
    /// <summary>
    /// A string representing the pictures Size. 
    /// For Example: 0,0
    /// </summary>
    public string Size
    {
        get
        {
            return this.GetXmlNodeString("v:fill/@size");
        }
        set
        {
            this.SetXmlNodeString("v:fill/@size", value, true);
        }
    }
    /// <summary>
    /// A string representing the pictures Origin
    /// </summary>
    public string Origin
    {
        get
        {
            return this.GetXmlNodeString("v:fill/@origin");
        }
        set
        {
            this.SetXmlNodeString("v:fill/@origin", value, true);
        }
    }
    /// <summary>
    /// A string representing the pictures position
    /// </summary>
    public string Position
    {
        get
        {
            return this.GetXmlNodeString("v:fill/@position");
        }
        set
        {
            this.SetXmlNodeString("v:fill/@position", value, true);
        }
    }
    /// <summary>
    /// The title for the fill
    /// </summary>
    public string Title
    {
        get
        {
            return this.GetXmlNodeString("v:fill/@o:title");
        }
        set
        {
            this.SetXmlNodeString("v:fill/@o:title", value, true);
        }
    }
    ExcelImage _image=null;
    /// <summary>
    /// The image is used when <see cref="ExcelVmlDrawingFill.Style"/> is set to  Pattern, Tile or Frame.
    /// </summary>
    public ExcelImage Image
    {
        get
        {
            if (this._image == null)
            {
                string? relId = this.RelId;
                this._image = new ExcelImage(this, new ePictureType[] { ePictureType.Svg, ePictureType.Ico, ePictureType.WebP });
                if (!string.IsNullOrEmpty(relId))
                {
                    this._image.ImageBytes = PictureStore.GetPicture(relId, this, out string contentType, out ePictureType pictureType);
                    this._image.Type = pictureType;
                }
            }
            return this._image;
        }
    }

    IPictureRelationDocument IPictureContainer.RelationDocument => this._fill._drawings.Worksheet.VmlDrawings;

    string IPictureContainer.ImageHash { get; set ; }
    Uri IPictureContainer.UriPic { get; set ; }
    ZipPackageRelationship IPictureContainer.RelPic { get; set; }
    void IPictureContainer.SetNewImage()
    {
        IPictureContainer? container = (IPictureContainer)this;
        //Create relationship
        this.SetXmlNodeString("v:fill/@o:relid", container.RelPic.Id);
    }
    void IPictureContainer.RemoveImage()
    {
        IPictureContainer? container = (IPictureContainer)this;
        IPictureRelationDocument? pictureRelationDocument = (IPictureRelationDocument)this._fill._drawings;
        pictureRelationDocument.Package.PictureStore.RemoveImage(container.ImageHash, this);
        pictureRelationDocument.RelatedPart.DeleteRelationship(container.RelPic.Id);
        pictureRelationDocument.Hashes.Remove(container.ImageHash);
    }

    internal string RelId 
    { 
        get
        {
            return this.GetXmlNodeString("v:fill/@o:relid");
        }
    }
}