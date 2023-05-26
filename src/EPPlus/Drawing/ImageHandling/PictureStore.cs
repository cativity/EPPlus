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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using OfficeOpenXml.Packaging;
using System.Linq;
namespace OfficeOpenXml.Drawing;

internal class ImageInfo
{
    internal string Hash { get; set; }
    internal Uri Uri { get; set; }
    internal int RefCount { get; set; }
    internal ZipPackagePart Part { get; set; }
    internal ExcelImageInfo Bounds { get; set; }
}
internal class PictureStore : IDisposable
{
    ExcelPackage _pck;
    internal static int _id = 1;
    internal Dictionary<string, ImageInfo> _images;
    public PictureStore(ExcelPackage pck)
    {
        this._pck = pck;
        this._images = this._pck.Workbook._images;
    }
    internal ImageInfo AddImage(byte[] image)
    {
        return this.AddImage(image, null, null);
    }
    internal ImageInfo AddImage(byte[] image, Uri uri, ePictureType? pictureType)
    {
        if (pictureType.HasValue == false)
        {
            pictureType = ePictureType.Jpg;
        }
#if (Core)
        SHA1? hashProvider = SHA1.Create();
#else
            var hashProvider = new SHA1CryptoServiceProvider();
#endif
        string? hash = BitConverter.ToString(hashProvider.ComputeHash(image)).Replace("-", "");
        lock (this._images)
        {
            if (this._images.ContainsKey(hash))
            {
                this._images[hash].RefCount++;
            }
            else
            {
                ZipPackagePart imagePart;
                string contentType;
                if (uri == null)
                {
                    string? extension = GetExtension(pictureType.Value);
                    contentType = GetContentType(extension);
                    uri = GetNewUri(this._pck.ZipPackage, "/xl/media/image{0}." + extension);
                    imagePart = this._pck.ZipPackage.CreatePart(uri, contentType, CompressionLevel.None, extension);
                    SaveImageToPart(image, imagePart);
                }
                else
                {
                    string? extension = GetExtension(uri);
                    contentType = GetContentType(extension);
                    pictureType = GetPictureType(extension);
                    if (this._pck.ZipPackage.PartExists(uri))
                    {
                        if(this._images.Values.Any(x=>x.Uri.OriginalString==uri.OriginalString))
                        {
                            uri = GetNewUri(this._pck.ZipPackage, "/xl/media/image{0}." + extension);
                            imagePart = this._pck.ZipPackage.CreatePart(uri, contentType, CompressionLevel.None, extension);
                            SaveImageToPart(image, imagePart);
                        }
                        else
                        {
                            imagePart = this._pck.ZipPackage.GetPart(uri);
                        }
                    }
                    else
                    {
                        imagePart = this._pck.ZipPackage.CreatePart(uri, contentType, CompressionLevel.None, extension);
                        SaveImageToPart(image, imagePart);
                    }
                }

                this._images.Add(hash,
                                 new ImageInfo()
                                 {
                                     Uri = uri,
                                     RefCount = 1,
                                     Hash = hash,
                                     Part = imagePart,
                                     Bounds = GetImageBounds(image, pictureType.Value, this._pck)
                                 });
            }
        }
        return this._images[hash];
    }

    private static void SaveImageToPart(byte[] image, ZipPackagePart imagePart)
    {
        Stream? stream = imagePart.GetStream(FileMode.Create, FileAccess.Write);
        stream.Write(image, 0, image.GetLength(0));
        stream.Flush();
    }

    internal static ExcelImageInfo GetImageBounds(byte[] image, ePictureType type, ExcelPackage pck)
    {
        ExcelImageInfo? ret = new ExcelImageInfo();
        MemoryStream? ms = RecyclableMemory.GetStream(image);
        ExcelImageSettings? s = pck.Settings.ImageSettings;

        if(s.GetImageBounds(ms, type, out double width, out double height, out double horizontalResolution, out double verticalResolution)==false)
        {
            throw (new InvalidOperationException($"No image handler for image type {type}"));
        }
        ret.Width = width;
        ret.Height = height;
        ret.HorizontalResolution = horizontalResolution;
        ret.VerticalResolution = verticalResolution;
        return ret;
    }
    internal static string GetExtension(Uri uri)
    {
        string? s = uri.OriginalString;
        int i = s.LastIndexOf('.');
        if(i>0)
        {
            return s.Substring(i + 1);
        }
        return null;
    }

    internal ImageInfo LoadImage(byte[] image, Uri uri, ZipPackagePart imagePart)
    {
#if (Core)
        SHA1? hashProvider = SHA1.Create();
#else
            var hashProvider = new SHA1CryptoServiceProvider();
#endif
        string? hash = BitConverter.ToString(hashProvider.ComputeHash(image)).Replace("-", "");
        if (this._images.ContainsKey(hash))
        {
            this._images[hash].RefCount++;
        }
        else
        {
            this._images.Add(hash, new ImageInfo() { Uri = uri, RefCount = 1, Hash = hash, Part = imagePart });
        }
        return this._images[hash];
    }
    internal void RemoveImage(string hash, IPictureContainer container)
    {
        lock (this._images)
        {
            if (this._images.ContainsKey(hash))
            {
                ImageInfo? ii = this._images[hash];
                ii.RefCount--;
                if (ii.RefCount == 0)
                {
                    this._pck.ZipPackage.DeletePart(ii.Uri);
                    this._images.Remove(hash);
                }
            }
            if(container.RelationDocument.Hashes.ContainsKey(hash))
            {
                container.RelationDocument.Hashes[hash].RefCount--;
                if (container.RelationDocument.Hashes[hash].RefCount <= 0)
                {
                    container.RelationDocument.Hashes.Remove(hash);
                }
                        
            }
        }
    }
    internal ImageInfo GetImageInfo(byte[] image)
    {
        string? hash = GetHash(image);
        if (this._images.ContainsKey(hash))
        {
            return this._images[hash];
        }
        else
        {
            return null;
        }
    }
    internal bool ImageExists(byte[] image)
    {
        string? hash = GetHash(image);
        return this._images.ContainsKey(hash);
    }

    internal static string GetHash(byte[] image)
    {
#if (Core)
        SHA1? hashProvider = SHA1.Create();
#else
            var hashProvider = new SHA1CryptoServiceProvider();
#endif
        return BitConverter.ToString(hashProvider.ComputeHash(image)).Replace("-", "");
    }

    private static Uri GetNewUri(ZipPackage package, string sUri)
    {
        Uri uri;
        do
        {
            uri = new Uri(string.Format(sUri, _id++), UriKind.Relative);
        }
        while (package.PartExists(uri));
        return uri;
    }
        
    internal static byte[] GetPicture(string relId, IPictureContainer container, out string contentType, out ePictureType pictureType)
    {
        container.RelPic = container.RelationDocument.RelatedPart.GetRelationship(relId);
        container.UriPic = UriHelper.ResolvePartUri(container.RelationDocument.RelatedUri, container.RelPic.TargetUri);
        ZipPackagePart part = container.RelationDocument.RelatedPart.Package.GetPart(container.UriPic);

        string? extension = Path.GetExtension(container.UriPic.OriginalString);
        contentType = GetContentType(extension);
        pictureType = GetPictureType(extension);
        return ((MemoryStream)part.GetStream()).ToArray();
    }
    internal static ePictureType GetPictureType(Uri uri)
    {
        string? ext = GetExtension(uri);
        return GetPictureType(ext);
    }
    internal static ePictureType GetPictureType(string extension)
    {
        if (extension.StartsWith(".", StringComparison.OrdinalIgnoreCase))
        {
            extension = extension.Substring(1);
        }

        switch (extension.ToLower(CultureInfo.InvariantCulture))
        {
            case "bmp":
            case "dib":
                return ePictureType.Bmp;
            case "jpg":
            case "jpeg":
            case "jfif":
            case "jpe":
            case "exif":
                return ePictureType.Jpg;
            case "gif":
                return ePictureType.Gif;
            case "png":
                return ePictureType.Png;
            case "emf":
                return ePictureType.Emf;
            case "emz":
                return ePictureType.Emz;
            case "tif":
            case "tiff":
                return ePictureType.Tif;
            case "wmf":
                return ePictureType.Wmf;
            case "wmz":
                return ePictureType.Wmz;
            case "webp":
                return ePictureType.WebP;
            case "ico":
                return ePictureType.Ico;
            case "svg":
                return ePictureType.Svg;
            default:
                throw (new InvalidOperationException($"Image with extension {extension} is not supported."));
        }
    }
    internal static string GetExtension(ePictureType type)
    {
        switch (type)
        {
            case ePictureType.Bmp:
                return "bmp";
            case ePictureType.Gif:
                return "gif";
            case ePictureType.Png:
                return "png";
            case ePictureType.Emf:
                return "emf";
            case ePictureType.Wmf:
                return "wmf";
            case ePictureType.Tif:
                return "tif";
            case ePictureType.WebP:
                return "webp";
            case ePictureType.Ico:
                return "ico";
            case ePictureType.Svg:
                return "svg";
            default:
                return "jpg";
        }
    }

    internal static string GetContentType(string extension)
    {
        if (extension.StartsWith(".", StringComparison.OrdinalIgnoreCase))
        {
            extension = extension.Substring(1);
        }

        switch (extension.ToLower(CultureInfo.InvariantCulture))
        {
            case "bmp":
            case "dib":
                return "image/bmp";
            case "jpg":
            case "jpeg":
            case "jfif":
            case "jpe":
                return "image/jpeg";
            case "gif":
                return "image/gif";
            case "png":
                return "image/png";
            case "cgm":
                return "image/cgm";
            case "emf":
            case "emz":
                return "image/x-emf";
            case "eps":
                return "image/x-eps";
            case "pcx":
                return "image/x-pcx";
            case "tga":
                return "image/x-tga";
            case "tif":
            case "tiff":
                return "image/x-tiff";
            case "wmf":
            case "wmz":
                return "image/x-wmf";
            case "svg":
                return "image/svg+xml";
            case "webp":
                return "image/webp";
            case "ico":
                return "image/x-icon";
            default:
                return "image/jpeg";
        }
    }
    internal static string SavePicture(byte[] image, IPictureContainer container, ePictureType type)
    {
        PictureStore? store = container.RelationDocument.Package.PictureStore;
        ImageInfo? ii = store.AddImage(image, container.UriPic, type);

        container.ImageHash = ii.Hash;
        Dictionary<string, HashInfo>? hashes = container.RelationDocument.Hashes;
        if (hashes.ContainsKey(ii.Hash))
        {
            string? relID = hashes[ii.Hash].RelId;
            container.RelPic = container.RelationDocument.RelatedPart.GetRelationship(relID);
            container.UriPic = UriHelper.ResolvePartUri(container.RelPic.SourceUri, container.RelPic.TargetUri);
            return relID;
        }
        else
        {
            container.UriPic = ii.Uri;
            container.ImageHash = ii.Hash;
        }

        //Set the Image and save it to the package.
        container.RelPic = container.RelationDocument.RelatedPart.CreateRelationship(UriHelper.GetRelativeUri(container.RelationDocument.RelatedUri, container.UriPic), TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");

        //AddNewPicture(img, picRelation.Id);
        hashes.Add(ii.Hash, new HashInfo(container.RelPic.Id) { RefCount = 1});

        return container.RelPic.Id;
    }

    public void Dispose()
    {
        this._images = null;
    }
}