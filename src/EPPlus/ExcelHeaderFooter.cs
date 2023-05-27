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
using System.Text.RegularExpressions;
using System.Drawing;
using System.Collections.Generic;
using OfficeOpenXml.Drawing.Vml;
using System.IO;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Packaging;

namespace OfficeOpenXml;

/// <summary>
/// How a picture will be aligned in the header/footer
/// </summary>
public enum PictureAlignment
{
    /// <summary>
    /// The picture will be added to the left aligned text
    /// </summary>
    Left,

    /// <summary>
    /// The picture will be added to the centered text
    /// </summary>
    Centered,

    /// <summary>
    /// The picture will be added to the right aligned text
    /// </summary>
    Right
}

#region class ExcelHeaderFooterText

/// <summary>
/// Print header and footer 
/// </summary>
public class ExcelHeaderFooterText
{
    ExcelWorksheet _ws;
    string _hf;

    internal ExcelHeaderFooterText(XmlNode TextNode, ExcelWorksheet ws, string hf)
    {
        this._ws = ws;
        this._hf = hf;

        if (TextNode == null || string.IsNullOrEmpty(TextNode.InnerText))
        {
            return;
        }

        string text = TextNode.InnerText;
        string code = text.Substring(0, 2);
        int startPos = 2;

        for (int pos = startPos; pos < text.Length - 2; pos++)
        {
            string newCode = text.Substring(pos, 2);

            if (newCode == "&C" || newCode == "&R")
            {
                this.SetText(code, text.Substring(startPos, pos - startPos));
                startPos = pos + 2;
                pos = startPos;
                code = newCode;
            }
        }

        this.SetText(code, text.Substring(startPos, text.Length - startPos));
    }

    private void SetText(string code, string text)
    {
        switch (code)
        {
            case "&L":
                this.LeftAlignedText = text;

                break;

            case "&C":
                this.CenteredText = text;

                break;

            default:
                this.RightAlignedText = text;

                break;
        }
    }

    /// <summary>
    /// Get/set the text to appear on the left hand side of the header (or footer) on the worksheet.
    /// </summary>
    public string LeftAlignedText = null;

    /// <summary>
    /// Get/set the text to appear in the center of the header (or footer) on the worksheet.
    /// </summary>
    public string CenteredText = null;

    /// <summary>
    /// Get/set the text to appear on the right hand side of the header (or footer) on the worksheet.
    /// </summary>
    public string RightAlignedText = null;

    /// <summary>
    /// Inserts a picture at the end of the text in the header or footer
    /// </summary>
    /// <param name="PictureFile">The image object containing the Picture</param>
    /// <param name="Alignment">Alignment. The image object will be inserted at the end of the Text.</param>
    public ExcelVmlDrawingPicture InsertPicture(FileInfo PictureFile, PictureAlignment Alignment)
    {
        string id = this.ValidateImage(Alignment);

        if (!PictureFile.Exists)
        {
            throw new FileNotFoundException(string.Format("{0} is missing", PictureFile.FullName));
        }

        Uri? uriPic = XmlHelper.GetNewUri(this._ws._package.ZipPackage,
                                          "/xl/media/"
                                          + PictureFile.Name.Substring(0, PictureFile.Name.Length - PictureFile.Extension.Length)
                                          + "{0}"
                                          + PictureFile.Extension);

        byte[]? imgBytes = File.ReadAllBytes(PictureFile.FullName);
        ImageInfo? ii = this._ws.Workbook._package.PictureStore.AddImage(imgBytes, uriPic, null);

        return this.AddImage(id, ii);
    }

    /// <summary>
    /// Inserts a picture at the end of the text in the header or footer
    /// </summary>
    /// <param name="PictureStream">The stream containing the picture</param>
    /// <param name="pictureType">The image format of the picture stream</param>
    /// <param name="Alignment">Alignment. The image object will be inserted at the end of the Text.</param>
    public ExcelVmlDrawingPicture InsertPicture(Stream PictureStream, ePictureType pictureType, PictureAlignment Alignment)
    {
        string id = this.ValidateImage(Alignment);

        byte[]? imgBytes = new byte[PictureStream.Length];
        _ = PictureStream.Seek(0, SeekOrigin.Begin);
        _ = PictureStream.Read(imgBytes, 0, imgBytes.Length);
        ImageInfo? ii = this._ws.Workbook._package.PictureStore.AddImage(imgBytes, null, pictureType);

        return this.AddImage(id, ii);
    }
#if NETFULL
        /// <summary>
        /// Inserts a picture at the end of the text in the header or footer
        /// </summary>
        /// <param name="Picture">The image object containing the Picture</param>
        /// <param name="Alignment">Alignment. The image object will be inserted at the end of the Text.</param>
        [Obsolete("This method is deprecated and is removed .NET standard/core. Please use overloads not referencing System.Drawing.Image")]
        public ExcelVmlDrawingPicture InsertPicture(Image Picture, PictureAlignment Alignment)
        {
            var b = ImageUtils.GetImageAsByteArray(Picture, out ePictureType type);
            using (var ms = new MemoryStream(b))
            {
                return InsertPicture(ms, type, Alignment);
            }
        }
#endif
    private ExcelVmlDrawingPicture AddImage(string id, ImageInfo ii)
    {
        double width = ii.Bounds.Width * 72 / ii.Bounds.HorizontalResolution, //Pixel --> Points
               height = ii.Bounds.Height * 72 / ii.Bounds.VerticalResolution; //Pixel --> Points

        //Add VML-drawing            
        return this._ws.HeaderFooter.Pictures.Add(id, ii.Uri, "", width, height);
    }

    private string ValidateImage(PictureAlignment Alignment)
    {
        string id = string.Concat(Alignment.ToString()[0], this._hf);

        foreach (ExcelVmlDrawingPicture image in this._ws.HeaderFooter.Pictures)
        {
            if (image.Id == id)
            {
                throw new InvalidOperationException("A picture already exists in this section");
            }
        }

        //Add the image placeholder to the end of the text
        switch (Alignment)
        {
            case PictureAlignment.Left:
                this.LeftAlignedText += ExcelHeaderFooter.Image;

                break;

            case PictureAlignment.Centered:
                this.CenteredText += ExcelHeaderFooter.Image;

                break;

            default:
                this.RightAlignedText += ExcelHeaderFooter.Image;

                break;
        }

        return id;
    }
}

#endregion

#region ExcelHeaderFooter

/// <summary>
/// Represents the Header and Footer on an Excel Worksheet
/// </summary>
public sealed class ExcelHeaderFooter : XmlHelper
{
    #region Static Properties

    /// <summary>
    /// The code for "current page #"
    /// </summary>
    public const string PageNumber = @"&P";

    /// <summary>
    /// The code for "total pages"
    /// </summary>
    public const string NumberOfPages = @"&N";

    /// <summary>
    /// The code for "text font color"
    /// RGB Color is specified as RRGGBB
    /// Theme Color is specified as TTSNN where TT is the theme color Id, S is either "+" or "-" of the tint/shade value, NN is the tint/shade value.
    /// </summary>
    public const string FontColor = @"&K";

    /// <summary>
    /// The code for "sheet tab name"
    /// </summary>
    public const string SheetName = @"&A";

    /// <summary>
    /// The code for "this workbook's file path"
    /// </summary>
    public const string FilePath = @"&Z";

    /// <summary>
    /// The code for "this workbook's file name"
    /// </summary>
    public const string FileName = @"&F";

    /// <summary>
    /// The code for "date"
    /// </summary>
    public const string CurrentDate = @"&D";

    /// <summary>
    /// The code for "time"
    /// </summary>
    public const string CurrentTime = @"&T";

    /// <summary>
    /// The code for "picture as background"
    /// </summary>
    public const string Image = @"&G";

    /// <summary>
    /// The code for "outline style"
    /// </summary>
    public const string OutlineStyle = @"&O";

    /// <summary>
    /// The code for "shadow style"
    /// </summary>
    public const string ShadowStyle = @"&H";

    #endregion

    #region ExcelHeaderFooter Private Properties

    internal ExcelHeaderFooterText _oddHeader;
    internal ExcelHeaderFooterText _oddFooter;
    internal ExcelHeaderFooterText _evenHeader;
    internal ExcelHeaderFooterText _evenFooter;
    internal ExcelHeaderFooterText _firstHeader;
    internal ExcelHeaderFooterText _firstFooter;
    private ExcelWorksheet _ws;

    #endregion

    #region ExcelHeaderFooter Constructor

    /// <summary>
    /// ExcelHeaderFooter Constructor
    /// </summary>
    /// <param name="nameSpaceManager"></param>
    /// <param name="topNode"></param>
    /// <param name="ws">The worksheet</param>
    internal ExcelHeaderFooter(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelWorksheet ws)
        : base(nameSpaceManager, topNode)
    {
        this._ws = ws;
        this.SchemaNodeOrder = new string[] { "headerFooter", "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter" };
    }

    #endregion

    #region alignWithMargins

    const string alignWithMarginsPath = "@alignWithMargins";

    /// <summary>
    /// Align with page margins
    /// </summary>
    public bool AlignWithMargins
    {
        get { return this.GetXmlNodeBool(alignWithMarginsPath); }
        set { this.SetXmlNodeString(alignWithMarginsPath, value ? "1" : "0"); }
    }

    #endregion

    #region differentOddEven

    const string differentOddEvenPath = "@differentOddEven";

    /// <summary>
    /// Displas different headers and footers on odd and even pages.
    /// </summary>
    public bool differentOddEven
    {
        get { return this.GetXmlNodeBool(differentOddEvenPath); }
        set { this.SetXmlNodeString(differentOddEvenPath, value ? "1" : "0"); }
    }

    #endregion

    #region differentFirst

    const string differentFirstPath = "@differentFirst";

    /// <summary>
    /// Display different headers and footers on the first page of the worksheet.
    /// </summary>
    public bool differentFirst
    {
        get { return this.GetXmlNodeBool(differentFirstPath); }
        set { this.SetXmlNodeString(differentFirstPath, value ? "1" : "0"); }
    }

    #endregion

    #region ScaleWithDoc

    const string scaleWithDocPath = "@scaleWithDoc";

    /// <summary>
    /// The header and footer should scale as you use the ShrinkToFit property on the document
    /// </summary>
    public bool ScaleWithDocument
    {
        get { return this.GetXmlNodeBool(scaleWithDocPath); }
        set { this.SetXmlNodeBool(scaleWithDocPath, value); }
    }

    #endregion

    #region ExcelHeaderFooter Public Properties

    /// <summary>
    /// Provides access to the header on odd numbered pages of the document.
    /// If you want the same header on both odd and even pages, then only set values in this ExcelHeaderFooterText class.
    /// </summary>
    public ExcelHeaderFooterText OddHeader
    {
        get
        {
            return this._oddHeader
                   ?? (this._oddHeader = new ExcelHeaderFooterText(this.TopNode.SelectSingleNode("d:oddHeader", this.NameSpaceManager), this._ws, "H"));
        }
    }

    /// <summary>
    /// Provides access to the footer on odd numbered pages of the document.
    /// If you want the same footer on both odd and even pages, then only set values in this ExcelHeaderFooterText class.
    /// </summary>
    public ExcelHeaderFooterText OddFooter
    {
        get
        {
            return this._oddFooter
                   ?? (this._oddFooter = new ExcelHeaderFooterText(this.TopNode.SelectSingleNode("d:oddFooter", this.NameSpaceManager), this._ws, "F"));
        }
    }

    // evenHeader and evenFooter set differentOddEven = true
    /// <summary>
    /// Provides access to the header on even numbered pages of the document.
    /// </summary>
    public ExcelHeaderFooterText EvenHeader
    {
        get
        {
            if (this._evenHeader == null)
            {
                this._evenHeader = new ExcelHeaderFooterText(this.TopNode.SelectSingleNode("d:evenHeader", this.NameSpaceManager), this._ws, "HEVEN");
                this.differentOddEven = true;
            }

            return this._evenHeader;
        }
    }

    /// <summary>
    /// Provides access to the footer on even numbered pages of the document.
    /// </summary>
    public ExcelHeaderFooterText EvenFooter
    {
        get
        {
            if (this._evenFooter == null)
            {
                this._evenFooter = new ExcelHeaderFooterText(this.TopNode.SelectSingleNode("d:evenFooter", this.NameSpaceManager), this._ws, "FEVEN");
                this.differentOddEven = true;
            }

            return this._evenFooter;
        }
    }

    /// <summary>
    /// Provides access to the header on the first page of the document.
    /// </summary>
    public ExcelHeaderFooterText FirstHeader
    {
        get
        {
            if (this._firstHeader == null)
            {
                this._firstHeader = new ExcelHeaderFooterText(this.TopNode.SelectSingleNode("d:firstHeader", this.NameSpaceManager), this._ws, "HFIRST");
                this.differentFirst = true;
            }

            return this._firstHeader;
        }
    }

    /// <summary>
    /// Provides access to the footer on the first page of the document.
    /// </summary>
    public ExcelHeaderFooterText FirstFooter
    {
        get
        {
            if (this._firstFooter == null)
            {
                this._firstFooter = new ExcelHeaderFooterText(this.TopNode.SelectSingleNode("d:firstFooter", this.NameSpaceManager), this._ws, "FFIRST");
                this.differentFirst = true;
            }

            return this._firstFooter;
        }
    }

    private ExcelVmlDrawingPictureCollection _vmlDrawingsHF = null;

    /// <summary>
    /// Vml drawings. Underlaying object for Header footer images
    /// </summary>
    public ExcelVmlDrawingPictureCollection Pictures
    {
        get
        {
            if (this._vmlDrawingsHF == null)
            {
                XmlNode? vmlNode = this._ws.WorksheetXml.SelectSingleNode("d:worksheet/d:legacyDrawingHF/@r:id", this.NameSpaceManager);

                if (vmlNode == null)
                {
                    this._vmlDrawingsHF = new ExcelVmlDrawingPictureCollection(this._ws, null);
                }
                else
                {
                    if (this._ws.Part.RelationshipExists(vmlNode.Value))
                    {
                        ZipPackageRelationship? rel = this._ws.Part.GetRelationship(vmlNode.Value);
                        Uri? vmlUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);

                        this._vmlDrawingsHF = new ExcelVmlDrawingPictureCollection(this._ws, vmlUri);
                        this._vmlDrawingsHF.RelId = rel.Id;
                    }
                }
            }

            return this._vmlDrawingsHF;
        }
    }

    #endregion

    #region Save //  ExcelHeaderFooter

    /// <summary>
    /// Saves the header and footer information to the worksheet XML
    /// </summary>
    internal void Save()
    {
        if (this._oddHeader != null)
        {
            this.SetXmlNodeStringPreserveWhiteSpace("d:oddHeader", GetText(this.OddHeader));
        }

        if (this._oddFooter != null)
        {
            this.SetXmlNodeStringPreserveWhiteSpace("d:oddFooter", GetText(this.OddFooter));
        }

        // only set evenHeader and evenFooter 
        if (this.differentOddEven)
        {
            if (this._evenHeader != null)
            {
                this.SetXmlNodeStringPreserveWhiteSpace("d:evenHeader", GetText(this.EvenHeader));
            }

            if (this._evenFooter != null)
            {
                this.SetXmlNodeStringPreserveWhiteSpace("d:evenFooter", GetText(this.EvenFooter));
            }
        }

        // only set firstHeader and firstFooter
        if (this.differentFirst)
        {
            if (this._firstHeader != null)
            {
                this.SetXmlNodeStringPreserveWhiteSpace("d:firstHeader", GetText(this.FirstHeader));
            }

            if (this._firstFooter != null)
            {
                this.SetXmlNodeStringPreserveWhiteSpace("d:firstFooter", GetText(this.FirstFooter));
            }
        }
    }

    internal void SaveHeaderFooterImages()
    {
        if (this._vmlDrawingsHF != null)
        {
            if (this._vmlDrawingsHF.Count == 0)
            {
                if (this._vmlDrawingsHF.Part != null)
                {
                    this._ws.Part.DeleteRelationship(this._vmlDrawingsHF.RelId);
                    this._ws._package.ZipPackage.DeletePart(this._vmlDrawingsHF.Uri);
                }
            }
            else
            {
                if (this._vmlDrawingsHF.Uri == null)
                {
                    this._vmlDrawingsHF.Uri = GetNewUri(this._ws._package.ZipPackage, @"/xl/drawings/vmlDrawing{0}.vml");
                }

                if (this._vmlDrawingsHF.Part == null)
                {
                    this._vmlDrawingsHF.Part = this._ws._package.ZipPackage.CreatePart(this._vmlDrawingsHF.Uri,
                                                                                       "application/vnd.openxmlformats-officedocument.vmlDrawing",
                                                                                       this._ws._package.Compression);

                    ZipPackageRelationship? rel = this._ws.Part.CreateRelationship(UriHelper.GetRelativeUri(this._ws.WorksheetUri, this._vmlDrawingsHF.Uri),
                                                                                   TargetMode.Internal,
                                                                                   ExcelPackage.schemaRelationships + "/vmlDrawing");

                    this._ws.SetHFLegacyDrawingRel(rel.Id);
                    this._vmlDrawingsHF.RelId = rel.Id;

                    foreach (ExcelVmlDrawingPicture draw in this._vmlDrawingsHF)
                    {
                        rel = this._vmlDrawingsHF.Part.CreateRelationship(UriHelper.GetRelativeUri(this._vmlDrawingsHF.Uri, draw.ImageUri),
                                                                          TargetMode.Internal,
                                                                          ExcelPackage.schemaRelationships + "/image");

                        draw.RelId = rel.Id;
                    }
                }
                else
                {
                    foreach (ExcelVmlDrawingPicture draw in this._vmlDrawingsHF)
                    {
                        if (string.IsNullOrEmpty(draw.RelId))
                        {
                            ZipPackageRelationship? rel =
                                this._vmlDrawingsHF.Part.CreateRelationship(UriHelper.GetRelativeUri(this._vmlDrawingsHF.Uri, draw.ImageUri),
                                                                            TargetMode.Internal,
                                                                            ExcelPackage.schemaRelationships + "/image");

                            draw.RelId = rel.Id;
                        }
                    }
                }

                this._vmlDrawingsHF.VmlDrawingXml.Save(this._vmlDrawingsHF.Part.GetStream());
            }
        }
    }

    private static string GetText(ExcelHeaderFooterText headerFooter)
    {
        string ret = "";

        if (headerFooter.LeftAlignedText != null)
        {
            ret += "&L" + headerFooter.LeftAlignedText;
        }

        if (headerFooter.CenteredText != null)
        {
            ret += "&C" + headerFooter.CenteredText;
        }

        if (headerFooter.RightAlignedText != null)
        {
            ret += "&R" + headerFooter.RightAlignedText;
        }

        return ret;
    }

    #endregion
}

#endregion