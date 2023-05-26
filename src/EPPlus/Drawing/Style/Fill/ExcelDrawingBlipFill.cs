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
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Packaging;
using System;
using System.Globalization;
using System.IO;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Fill;

/// <summary>
/// A picture fill for a drawing
/// </summary>
public class ExcelDrawingBlipFill : ExcelDrawingFillBase, IPictureContainer
{
    string[] _schemaNodeOrder;
    private readonly IPictureRelationDocument _pictureRelationDocument;
    internal ExcelDrawingBlipFill(IPictureRelationDocument pictureRelationDocument, XmlNamespaceManager nsm, XmlNode topNode, string fillPath, string[] schemaNodeOrder, Action initXml) : base(nsm, topNode, fillPath, initXml)
    {
        this._schemaNodeOrder = schemaNodeOrder;
        this._pictureRelationDocument = pictureRelationDocument;
        this.Image = new ExcelImage(this);
        this.GetXml();
    }
    /// <summary>
    /// The image used in the fill operation.
    /// </summary>
    public ExcelImage Image { get; }
    /// <summary>
    /// The image should be stretched to fill the target.
    /// </summary>
    public bool Stretch { get; set; } = false;
    /// <summary>
    /// Offset in percentage from the edge of the shapes bounding box. This property only apply when Stretch is set to true.        
    /// <seealso cref="Stretch"/>
    /// </summary>
    public ExcelDrawingRectangle StretchOffset { get; private set; } = new ExcelDrawingRectangle(0);
    /// <summary>
    /// The portion of the image to be used for the fill.
    /// Offset values are in percentage from the borders of the image
    /// </summary>
    public ExcelDrawingRectangle SourceRectangle { get; private set; } = new ExcelDrawingRectangle(0);
    /// <summary>
    /// The image should be tiled to fill the available space
    /// </summary>
    public ExcelDrawingBlipFillTile Tile
    {
        get;
        private set;
    } = new ExcelDrawingBlipFillTile();
    /// <summary>
    /// The type of fill
    /// </summary>
    public override eFillStyle Style
    {
        get
        {
            return eFillStyle.BlipFill;
        }
    }
    ExcelDrawingBlipEffects _effects=null;
    /// <summary>
    /// Blip fill effects
    /// </summary>
    public ExcelDrawingBlipEffects Effects
    {
        get { return this._effects ?? (this._effects = new ExcelDrawingBlipEffects(this._nsm, this._topNode.SelectSingleNode("a:blip", this._nsm))); }
    }
    internal override string NodeName
    {
        get
        {
            return "a:blipFill";
        }
    }

    internal override void GetXml()
    {
        string? relId = this._xml.GetXmlNodeString("a:blip/@r:embed");
        if (!string.IsNullOrEmpty(relId))
        {
            byte[]? img = PictureStore.GetPicture(relId, this, out string contentType, out ePictureType pictureType);
            this.Image.Type = pictureType;
            this.Image.ImageBytes = img;
            this.ContentType = contentType;
        }

        this.SourceRectangle = new ExcelDrawingRectangle(this._xml, "a:srcRect/", 0);
        this.Stretch = this._xml.ExistsNode("a:stretch");
        if (this.Stretch)
        {
            this.StretchOffset = new ExcelDrawingRectangle(this._xml, "a:stretch/a:fillRect/", 0);
        }

        this.Tile = new ExcelDrawingBlipFillTile(this._xml);
    }

    internal override void SetXml(XmlNamespaceManager nsm, XmlNode node)
    {
        this._initXml?.Invoke();
        if (this._xml == null)
        {
            this.InitXml(nsm, node.FirstChild, "");
        }

        this.CheckTypeChange(this.NodeName);

        if (this.SourceRectangle.BottomOffset != 0)
        {
            this._xml.SetXmlNodePercentage("a:srcRect/@b", this.SourceRectangle.BottomOffset);
        }

        if (this.SourceRectangle.TopOffset != 0)
        {
            this._xml.SetXmlNodePercentage("a:srcRect/@t", this.SourceRectangle.TopOffset);
        }

        if (this.SourceRectangle.LeftOffset != 0)
        {
            this._xml.SetXmlNodePercentage("a:srcRect/@l", this.SourceRectangle.LeftOffset);
        }

        if (this.SourceRectangle.RightOffset != 0)
        {
            this._xml.SetXmlNodePercentage("a:srcRect/@r", this.SourceRectangle.RightOffset);
        }

        if (this.Tile.Alignment != null && this.Tile.FlipMode != null)
        {
            if (this.Tile.Alignment.HasValue)
            {
                this._xml.SetXmlNodeString("a:tile/@algn", this.Tile.Alignment.Value.TranslateString());
            }

            if (this.Tile.FlipMode.HasValue)
            {
                this._xml.SetXmlNodeString("a:tile/@flip", this.Tile.FlipMode.Value.ToString().ToLower());
            }

            this._xml.SetXmlNodePercentage("a:tile/@sx", this.Tile.HorizontalRatio, false);
            this._xml.SetXmlNodePercentage("a:tile/@sy", this.Tile.VerticalRatio, false);
            this._xml.SetXmlNodeString("a:tile/@tx", (this.Tile.HorizontalOffset * ExcelDrawing.EMU_PER_PIXEL).ToString(CultureInfo.InvariantCulture));
            this._xml.SetXmlNodeString("a:tile/@ty", (this.Tile.VerticalOffset * ExcelDrawing.EMU_PER_PIXEL).ToString(CultureInfo.InvariantCulture));
        }

        if (this.Stretch)
        {
            this._xml.SetXmlNodePercentage("a:stretch/a:fillRect/@b", this.StretchOffset.BottomOffset);
            this._xml.SetXmlNodePercentage("a:stretch/a:fillRect/@t", this.StretchOffset.TopOffset);
            this._xml.SetXmlNodePercentage("a:stretch/a:fillRect/@l", this.StretchOffset.LeftOffset);
            this._xml.SetXmlNodePercentage("a:stretch/a:fillRect/@r", this.StretchOffset.RightOffset);
        }
    }

    internal override void UpdateXml()
    {
        this.SetXml(this._xml.NameSpaceManager, this._xml.TopNode);
    }

    internal void AddImage(FileInfo file)
    {
        if (!file.Exists)
        {
            throw (new ArgumentException($"File {file.FullName} does not exist."));
        }
        byte[]? img = File.ReadAllBytes(file.FullName);
        string? extension = file.Extension;
        this.ContentType = PictureStore.GetContentType(extension);
        this.Image.SetImage(img, PictureStore.GetPictureType(extension));
    }
    #region IPictureContainer

    string IPictureContainer.ImageHash
    {
        get;
        set;
    }
    Uri IPictureContainer.UriPic
    {
        get;
        set;
    }
    ZipPackageRelationship IPictureContainer.RelPic
    {
        get;
        set;
    }
    void IPictureContainer.SetNewImage()
    {
        IPictureContainer container = this;
        //Create relationship
        this._xml.SetXmlNodeString("a:blip/@r:embed", container.RelPic.Id);
    }
    void IPictureContainer.RemoveImage()
    {
        IPictureContainer container = this;
        this._pictureRelationDocument.Package.PictureStore.RemoveImage(container.ImageHash, this);
        this._pictureRelationDocument.RelatedPart.DeleteRelationship(container.RelPic.Id);
        this._pictureRelationDocument.Hashes.Remove(container.ImageHash);
    }
    internal string ContentType
    {
        get;
        set;
    }

    IPictureRelationDocument IPictureContainer.RelationDocument { get => this._pictureRelationDocument; }
    #endregion
}