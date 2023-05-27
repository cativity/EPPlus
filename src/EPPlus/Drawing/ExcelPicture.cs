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
using System.Text;
using System.Xml;
using System.IO;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Packaging;
#if NETFULL
using System.Drawing.Imaging;
#endif
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// An image object
    /// </summary>
    public sealed class ExcelPicture : ExcelDrawing, IPictureContainer
    {
        #region "Constructors"

        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, Uri hyperlink, ePictureType type)
            : base(drawings, node, "xdr:pic", "xdr:nvPicPr/xdr:cNvPr")
        {
            this.CreatePicNode(node, type);
            this.Hyperlink = hyperlink;
            this.Image = new ExcelImage(this);
        }

        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, ExcelGroupShape shape = null)
            : base(drawings, node, "xdr:pic", "xdr:nvPicPr/xdr:cNvPr", shape)
        {
            XmlNode picNode = node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip", drawings.NameSpaceManager);

            if (picNode != null && picNode.Attributes["embed", ExcelPackage.schemaRelationships] != null)
            {
                IPictureContainer container = this;
                container.RelPic = drawings.Part.GetRelationship(picNode.Attributes["embed", ExcelPackage.schemaRelationships].Value);
                container.UriPic = UriHelper.ResolvePartUri(drawings.UriDrawing, container.RelPic.TargetUri);

                string? extension = Path.GetExtension(container.UriPic.OriginalString);
                this.ContentType = PictureStore.GetContentType(extension);

                if (drawings.Part.Package.PartExists(container.UriPic))
                {
                    this.Part = drawings.Part.Package.GetPart(container.UriPic);
                }
                else
                {
                    this.Part = null;

                    return;
                }

                byte[] iby = ((MemoryStream)this.Part.GetStream()).ToArray();
                this.Image = new ExcelImage(this);
                this.Image.Type = PictureStore.GetPictureType(extension);
                this.Image.ImageBytes = iby;
                ImageInfo? ii = this._drawings._package.PictureStore.LoadImage(iby, container.UriPic, this.Part);
                IPictureRelationDocument? pd = (IPictureRelationDocument)this._drawings;

                if (pd.Hashes.ContainsKey(ii.Hash))
                {
                    pd.Hashes[ii.Hash].RefCount++;
                }
                else
                {
                    pd.Hashes.Add(ii.Hash, new HashInfo(container.RelPic.Id) { RefCount = 1 });
                }

                container.ImageHash = ii.Hash;
            }
        }

        private void SetRelId(XmlNode node, ePictureType type, string relID)
        {
            //Create relationship
            node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", this.NameSpaceManager).Value = relID;

            if (type == ePictureType.Svg)
            {
                node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/a:extLst/a:ext/asvg:svgBlip/@r:embed", this.NameSpaceManager).Value = relID;
            }
        }

        /// <summary>
        /// The type of drawing
        /// </summary>
        public override eDrawingType DrawingType
        {
            get { return eDrawingType.Picture; }
        }
#if !NET35 && !NET40
        internal async Task LoadImageAsync(Stream stream, ePictureType type)
        {
            byte[]? img = new byte[stream.Length];
            stream.Seek(0, SeekOrigin.Begin);
            await stream.ReadAsync(img, 0, (int)stream.Length).ConfigureAwait(false);

            this.SaveImageToPackage(type, img);
        }
#endif
        internal void LoadImage(Stream stream, ePictureType type)
        {
            byte[]? img = new byte[stream.Length];
            stream.Seek(0, SeekOrigin.Begin);
            stream.Read(img, 0, (int)stream.Length);

            this.SaveImageToPackage(type, img);
        }

        private void SaveImageToPackage(ePictureType type, byte[] img)
        {
            ZipPackage? package = this._drawings.Worksheet._package.ZipPackage;

            if (type == ePictureType.Emz || type == ePictureType.Wmz)
            {
                img = ImageReader.ExtractImage(img, out ePictureType? pt);

                if (pt == null)
                {
                    throw new InvalidDataException($"Invalid image of type {type}");
                }

                type = pt.Value;
            }

            this.ContentType = PictureStore.GetContentType(type.ToString());
            Uri? newUri = GetNewUri(package, "/xl/media/image{0}." + type.ToString());
            PictureStore? store = this._drawings._package.PictureStore;
            IPictureRelationDocument? pc = this._drawings as IPictureRelationDocument;
            ImageInfo? ii = store.AddImage(img, newUri, type);

            IPictureContainer container = this;
            container.UriPic = ii.Uri;
            string relId;

            if (!pc.Hashes.ContainsKey(ii.Hash))
            {
                this.Part = ii.Part;

                container.RelPic = this._drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(this._drawings.UriDrawing, ii.Uri),
                                                                          TargetMode.Internal,
                                                                          ExcelPackage.schemaRelationships + "/image");

                relId = container.RelPic.Id;
                pc.Hashes.Add(ii.Hash, new HashInfo(relId));
                AddNewPicture(img, relId);
            }
            else
            {
                relId = pc.Hashes[ii.Hash].RelId;
                ZipPackageRelationship? rel = this._drawings.Part.GetRelationship(relId);
                container.UriPic = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            }

            container.ImageHash = ii.Hash;

            using (MemoryStream? ms = RecyclableMemory.GetStream(img))
            {
                this.Image.Bounds = PictureStore.GetImageBounds(img, type, this._drawings._package);
                this.Image.ImageBytes = img;
                this.Image.Type = type;
                double width = this.Image.Bounds.Width / (this.Image.Bounds.HorizontalResolution / STANDARD_DPI);
                double height = this.Image.Bounds.Height / (this.Image.Bounds.VerticalResolution / STANDARD_DPI);
                this.SetPosDefaults((float)width, (float)height);
            }

            //Create relationship
            this.SetRelId(this.TopNode, type, relId);

            //TopNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relId;
            ZipPackage.Flush();
        }

        private void CreatePicNode(XmlNode node, ePictureType type)
        {
            XmlNode? picNode = this.CreateNode("xdr:pic");
            picNode.InnerXml = this.PicStartXml(type);

            node.InsertAfter(node.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings), picNode);
        }

        private static void AddNewPicture(byte[] img, string relID)
        {
            ExcelDrawings.ImageCompare? newPic = new ExcelDrawings.ImageCompare();
            newPic.image = img;
            newPic.relID = relID;

            //_drawings._pics.Add(newPic);
        }

        #endregion

        private void SetPosDefaults(float width, float height)
        {
            this.EditAs = eEditAs.OneCell;
            this.SetPixelWidth(width);
            this.SetPixelHeight(height);
            this._width = this.GetPixelWidth();
            this._height = this.GetPixelHeight();
        }

        private string PicStartXml(ePictureType type)
        {
            StringBuilder xml = new StringBuilder();

            xml.Append("<xdr:nvPicPr>");
            xml.AppendFormat("<xdr:cNvPr id=\"{0}\" descr=\"\" />", this._id);
            xml.Append("<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill>");

            if (type == ePictureType.Svg)
            {
                xml.Append("<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"\" cstate=\"print\"><a:extLst><a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\"><a14:useLocalDpi xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/></a:ext><a:ext uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\"><asvg:svgBlip xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" r:embed=\"\"/></a:ext></a:extLst></a:blip>");
            }
            else
            {
                xml.Append("<a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"\" cstate=\"print\" />");
            }

            xml.Append("<a:stretch><a:fillRect /> </a:stretch> </xdr:blipFill> <xdr:spPr> <a:xfrm> <a:off x=\"0\" y=\"0\" />  <a:ext cx=\"0\" cy=\"0\" /> </a:xfrm> <a:prstGeom prst=\"rect\"> <a:avLst /> </a:prstGeom> </xdr:spPr>");

            return xml.ToString();
        }

        /// <summary>
        /// The image
        /// </summary>
        public ExcelImage Image { get; }

        internal string ContentType { get; set; }

        /// <summary>
        /// Set the size of the image in percent from the orginal size
        /// Note that resizing columns / rows after using this function will effect the size of the picture
        /// </summary>
        /// <param name="Percent">Percent</param>
        public override void SetSize(int Percent)
        {
            if (this.Image.ImageBytes == null)
            {
                base.SetSize(Percent);
            }
            else
            {
                this._width = this.Image.Bounds.Width / (this.Image.Bounds.HorizontalResolution / STANDARD_DPI);
                this._height = this.Image.Bounds.Height / (this.Image.Bounds.VerticalResolution / STANDARD_DPI);

                this._width = (int)(this._width * ((double)Percent / 100));
                this._height = (int)(this._height * ((double)Percent / 100));

                this._doNotAdjust = true;
                this.SetPixelWidth(this._width);
                this.SetPixelHeight(this._height);
                this._doNotAdjust = false;
            }
        }

        internal ZipPackagePart Part;

        internal new string Id
        {
            get { return this.Name; }
        }

        ExcelDrawingFill _fill = null;

        /// <summary>
        /// Access to Fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get { return this._fill ??= new ExcelDrawingFill(this._drawings, this.NameSpaceManager, this.TopNode, "xdr:pic/xdr:spPr", this.SchemaNodeOrder); }
        }

        ExcelDrawingBorder _border = null;

        /// <summary>
        /// Access to Fill properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                return this._border ??= new ExcelDrawingBorder(this._drawings,
                                                               this.NameSpaceManager,
                                                               this.TopNode,
                                                               "xdr:pic/xdr:spPr/a:ln",
                                                               this.SchemaNodeOrder);
            }
        }

        ExcelDrawingEffectStyle _effect = null;

        /// <summary>
        /// Effects
        /// </summary>
        public ExcelDrawingEffectStyle Effect
        {
            get
            {
                return this._effect ??= new ExcelDrawingEffectStyle(this._drawings,
                                                                    this.NameSpaceManager,
                                                                    this.TopNode,
                                                                    "xdr:pic/xdr:spPr/a:effectLst",
                                                                    this.SchemaNodeOrder);
            }
        }

        const string _preferRelativeResizePath = "xdr:pic/xdr:nvPicPr/xdr:cNvPicPr/@preferRelativeResize";

        /// <summary>
        /// Relative to original picture size
        /// </summary>
        public bool PreferRelativeResize
        {
            get { return this.GetXmlNodeBool(_preferRelativeResizePath); }
            set { this.SetXmlNodeBool(_preferRelativeResizePath, value); }
        }

        const string _lockAspectRatioPath = "xdr:pic/xdr:nvPicPr/xdr:cNvPicPr/a:picLocks/@noChangeAspect";

        /// <summary>
        /// Lock aspect ratio
        /// </summary>
        public bool LockAspectRatio
        {
            get { return this.GetXmlNodeBool(_lockAspectRatioPath); }
            set { this.SetXmlNodeBool(_lockAspectRatioPath, value); }
        }

        internal override void CellAnchorChanged()
        {
            base.CellAnchorChanged();

            if (this._fill != null)
            {
                this._fill.SetTopNode(this.TopNode);
            }

            if (this._border != null)
            {
                this._border.TopNode = this.TopNode;
            }

            if (this._effect != null)
            {
                this._effect.TopNode = this.TopNode;
            }
        }

        internal override void DeleteMe()
        {
            IPictureContainer container = this;
            this._drawings._package.PictureStore.RemoveImage(container.ImageHash, this);
            base.DeleteMe();
        }

        /// <summary>
        /// Dispose the object
        /// </summary>
        public override void Dispose()
        {
            //base.Dispose();
            //Hyperlink = null;
            //_image.Dispose();
            //_image = null;            
        }

        void IPictureContainer.RemoveImage()
        {
            IPictureContainer container = this;
            IPictureRelationDocument? relDoc = (IPictureRelationDocument)this._drawings;

            if (relDoc.Hashes.TryGetValue(container.ImageHash, out HashInfo hi))
            {
                if (hi.RefCount <= 1)
                {
                    relDoc.Package.PictureStore.RemoveImage(container.ImageHash, this);
                    relDoc.RelatedPart.DeleteRelationship(container.RelPic.Id);
                    relDoc.Hashes.Remove(container.ImageHash);
                }
                else
                {
                    hi.RefCount--;
                }
            }
        }

        void IPictureContainer.SetNewImage()
        {
            string? relId = ((IPictureContainer)this).RelPic.Id;
            this.TopNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", this.NameSpaceManager).Value = relId;

            if (this.Image.Type == ePictureType.Svg)
            {
                this.TopNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/a:extLst/a:ext/asvg:svgBlip/@r:embed", this.NameSpaceManager).Value = relId;
            }
        }

        string IPictureContainer.ImageHash { get; set; }

        Uri IPictureContainer.UriPic { get; set; }

        ZipPackageRelationship IPictureContainer.RelPic { get; set; }

        IPictureRelationDocument IPictureContainer.RelationDocument => this._drawings;
    }
}