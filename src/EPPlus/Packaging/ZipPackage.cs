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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Constants;

namespace OfficeOpenXml.Packaging;

/// <summary>
/// Represent an OOXML Zip package.
/// </summary>
internal partial class ZipPackage : ZipPackagePartBase, IDisposable
{
    internal class ContentType
    {
        internal string Name;
        internal bool IsExtension;
        internal string Match;

        public ContentType(string name, bool isExtension, string match)
        {
            this.Name = name;
            this.IsExtension = isExtension;
            this.Match = match;
        }
    }

    Dictionary<string, ZipPackagePart> Parts = new Dictionary<string, ZipPackagePart>(StringComparer.OrdinalIgnoreCase);
    internal Dictionary<string, ContentType> _contentTypes = new Dictionary<string, ContentType>(StringComparer.OrdinalIgnoreCase);
    internal char _dirSeparator = '0';

    internal ZipPackage() => this.AddNew();

    private void AddNew()
    {
        this._contentTypes.Add("xml", new ContentType(ExcelPackage.schemaXmlExtension, true, "xml"));
        this._contentTypes.Add("rels", new ContentType(ExcelPackage.schemaRelsExtension, true, "rels"));
    }

    internal ZipInputStream _zip;

    internal ZipPackage(Stream stream)
    {
        bool hasContentTypeXml = false;

        if (stream == null || stream.Length == 0)
        {
            this.AddNew();
        }
        else
        {
            Dictionary<string, string>? rels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _ = stream.Seek(0, SeekOrigin.Begin);

            //using (ZipInputStream zip = new ZipInputStream(stream))
            //{
            this._zip = new ZipInputStream(stream);
            ZipEntry? e = this._zip.GetNextEntry();

            if (e == null)
            {
                throw
                    new InvalidDataException("The file is not a valid Package file. If the file is encrypted, please supply the password in the constructor.");
            }

            while (e != null)
            {
                this.GetDirSeparator(e);

                if (e.UncompressedSize > 0)
                {
                    if (e.FileName.Equals("[content_types].xml", StringComparison.OrdinalIgnoreCase))
                    {
                        this.AddContentTypes(Encoding.UTF8.GetString(GetZipEntryAsByteArray(this._zip, e)));
                        hasContentTypeXml = true;
                    }

                    else if (e.FileName.Equals($"_rels{this._dirSeparator}.rels", StringComparison.OrdinalIgnoreCase))
                    {
                        this.ReadRelation(Encoding.UTF8.GetString(GetZipEntryAsByteArray(this._zip, e)), "");
                    }
                    else if (e.FileName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                    {
                        byte[]? ba = GetZipEntryAsByteArray(this._zip, e);
                        rels.Add(GetUriKey(e.FileName), Encoding.UTF8.GetString(ba));
                    }
                    else
                    {
                        this.ExtractEntryToPart(this._zip, e);
                    }
                }

                e = this._zip.GetNextEntry();
            }

            if (this._dirSeparator == '0')
            {
                this._dirSeparator = '/';
            }

            foreach (KeyValuePair<string, ZipPackagePart> p in this.Parts)
            {
                string name = Path.GetFileName(p.Key);
                string extension = Path.GetExtension(p.Key);
                string relFile = string.Format("{0}_rels/{1}.rels", p.Key.Substring(0, p.Key.Length - name.Length), name);

                if (rels.ContainsKey(relFile))
                {
                    p.Value.ReadRelation(rels[relFile], p.Value.Uri.OriginalString);
                }

                if (this._contentTypes.ContainsKey(p.Key))
                {
                    p.Value.ContentType = this._contentTypes[p.Key].Name;
                }
                else if (extension.Length > 1 && this._contentTypes.ContainsKey(extension.Substring(1)))
                {
                    p.Value.ContentType = this._contentTypes[extension.Substring(1)].Name;
                }
            }

            if (!hasContentTypeXml)
            {
                throw
                    new InvalidDataException("The file is not a valid Package file. If the file is encrypted, please supply the password in the constructor.");
            }

            if (!hasContentTypeXml)
            {
                throw
                    new InvalidDataException("The file is not a valid Package file. If the file is encrypted, please supply the password in the constructor.");
            }

            //zip.Close();
            //zip.Dispose();
            //}
        }
    }

    private static byte[] GetZipEntryAsByteArray(ZipInputStream zip, ZipEntry e)
    {
        byte[]? b = new byte[e.UncompressedSize];
        _ = zip.Read(b, 0, (int)e.UncompressedSize);

        return b;
    }

    private void ExtractEntryToPart(ZipInputStream zip, ZipEntry e)
    {
        ZipPackagePart? part = new ZipPackagePart(this, e);

        long rest = e.UncompressedSize;

        if (rest > int.MaxValue)
        {
            part.Stream = null; //Over 2GB, we use the zip stream directly instead.
        }
        else
        {
            const int BATCH_SIZE = 0x100000;
            part.Stream = RecyclableMemory.GetStream();

            while (rest > 0)
            {
                int bufferSize = rest > BATCH_SIZE ? BATCH_SIZE : (int)rest;
                byte[]? b = new byte[bufferSize];
                int size = zip.Read(b, 0, bufferSize);
                part.Stream.Write(b, 0, size);
                rest -= size;
            }
        }

        this.Parts.Add(GetUriKey(e.FileName), part);
    }

    private void GetDirSeparator(ZipEntry e)
    {
        if (this._dirSeparator == '0')
        {
            if (e.FileName.Contains("\\"))
            {
                this._dirSeparator = '\\';
            }
            else if (e.FileName.Contains("/"))
            {
                this._dirSeparator = '/';
            }
        }
    }

    private void AddContentTypes(string xml)
    {
        XmlDocument? doc = new XmlDocument();
        XmlHelper.LoadXmlSafe(doc, xml, Encoding.UTF8);

        foreach (XmlElement c in doc.DocumentElement.ChildNodes)
        {
            ContentType ct;

            if (string.IsNullOrEmpty(c.GetAttribute("Extension")))
            {
                ct = new ContentType(c.GetAttribute("ContentType"), false, c.GetAttribute("PartName"));
                this._contentTypes.Add(GetUriKey(ct.Match), ct);
            }
            else
            {
                ct = new ContentType(c.GetAttribute("ContentType"), true, c.GetAttribute("Extension"));
                this._contentTypes.Add(ct.Match, ct);
            }
        }
    }

    #region Methods

    internal ZipPackagePart CreatePart(Uri partUri, string contentType) => this.CreatePart(partUri, contentType, CompressionLevel.Default);

    internal ZipPackagePart CreatePart(Uri partUri, string contentType, CompressionLevel compressionLevel, string extension = null)
    {
        if (this.PartExists(partUri))
        {
            throw new InvalidOperationException("Part already exist");
        }

        ZipPackagePart? part = new ZipPackagePart(this, partUri, contentType, compressionLevel);

        if (string.IsNullOrEmpty(extension))
        {
            this._contentTypes.Add(GetUriKey(part.Uri.OriginalString), new ContentType(contentType, false, part.Uri.OriginalString));
        }
        else
        {
            if (!this._contentTypes.ContainsKey(extension))
            {
                this._contentTypes.Add(extension, new ContentType(contentType, true, extension));
            }
        }

        this.Parts.Add(GetUriKey(part.Uri.OriginalString), part);

        return part;
    }

    internal ZipPackagePart CreatePart(Uri partUri, ZipPackagePart sourcePart)
    {
        ZipPackagePart? destPart = this.CreatePart(partUri, sourcePart.ContentType);
        Stream? destStream = destPart.GetStream(FileMode.Create, FileAccess.Write);
        Stream? sourceStream = sourcePart.GetStream();
        byte[]? b = ((MemoryStream)sourceStream).GetBuffer();
        destStream.Write(b, 0, b.Length);
        destStream.Flush();

        return destPart;
    }

    internal ZipPackagePart CreatePart(Uri partUri, string contentType, string xml)
    {
        ZipPackagePart? destPart = this.CreatePart(partUri, contentType);
        StreamWriter? destStream = new StreamWriter(destPart.GetStream(FileMode.Create, FileAccess.Write));
        destStream.Write(xml);
        destStream.Flush();

        return destPart;
    }

    internal ZipPackagePart GetPart(Uri partUri)
    {
        if (this.PartExists(partUri))
        {
            return this.Parts.Single(x => x.Key.Equals(GetUriKey(partUri.OriginalString), StringComparison.OrdinalIgnoreCase)).Value;
        }
        else
        {
            throw new InvalidOperationException("Part does not exist.");
        }
    }

    internal static string GetUriKey(string uri)
    {
        string ret = uri.Replace('\\', '/');

        if (ret[0] != '/')
        {
            ret = '/' + ret;
        }

        return ret;
    }

    internal bool PartExists(Uri partUri)
    {
        string? uriKey = GetUriKey(partUri.OriginalString.ToLowerInvariant());

        return this.Parts.ContainsKey(uriKey);
    }

    #endregion

    internal void DeletePart(Uri Uri)
    {
        List<object[]>? delList = new List<object[]>();

        foreach (ZipPackagePart? p in this.Parts.Values)
        {
            foreach (ZipPackageRelationship? r in p.GetRelationships())
            {
                if (r.TargetUri != null
                    && UriHelper.ResolvePartUri(p.Uri, r.TargetUri).OriginalString.Equals(Uri.OriginalString, StringComparison.OrdinalIgnoreCase))
                {
                    delList.Add(new object[] { r.Id, p });
                }
            }
        }

        foreach (object[]? o in delList)
        {
            ((ZipPackagePart)o[1]).DeleteRelationship(o[0].ToString());
        }

        ZipPackageRelationshipCollection? rels = this.GetPart(Uri).GetRelationships();

        while (rels.Count > 0)
        {
            rels.Remove(rels.First().Id);
        }

        _ = this._contentTypes.Remove(GetUriKey(Uri.OriginalString));

        //remove all relations
        _ = this.Parts.Remove(GetUriKey(Uri.OriginalString));
    }

    internal void Save(Stream stream)
    {
        Encoding? enc = Encoding.UTF8;
        ZipOutputStream os = new ZipOutputStream(stream, true);
        os.EnableZip64 = Zip64Option.AsNecessary;
        os.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)this._compression;

        /**** ContentType****/
        _ = os.PutNextEntry("[Content_Types].xml");
        byte[] b = enc.GetBytes(this.GetContentTypeXml());
        os.Write(b, 0, b.Length);
        /**** Top Rels ****/
        this._rels.WriteZip(os, $"_rels/.rels");
        ZipPackagePart ssPart = null;

        foreach (ZipPackagePart? part in this.Parts.Values)
        {
            if (part.ContentType != ContentTypes.contentTypeSharedString)
            {
                part.WriteZip(os);
            }
            else
            {
                ssPart = part;
            }
        }

        //Shared strings must be saved after all worksheets. The ss dictionary is populated when that workheets are saved (to get the best performance).
        if (ssPart != null)
        {
            ssPart.WriteZip(os);
        }

        os.Flush();

        os.Close();
        os.Dispose();

        //return ms;
    }

    private string GetContentTypeXml()
    {
        StringBuilder xml =
            new
                StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");

        foreach (ContentType ct in this._contentTypes.Values)
        {
            if (ct.IsExtension)
            {
                _ = xml.AppendFormat("<Default ContentType=\"{0}\" Extension=\"{1}\"/>", ct.Name, ct.Match);
            }
            else
            {
                _ = xml.AppendFormat("<Override ContentType=\"{0}\" PartName=\"{1}\" />", ct.Name, GetUriKey(ct.Match));
            }
        }

        _ = xml.Append("</Types>");

        return xml.ToString();
    }

    internal static void Flush()
    {
    }

    internal static void Close()
    {
    }

    public void Dispose()
    {
        foreach (ZipPackagePart? part in this.Parts.Values)
        {
            part.Dispose();
        }

        this._zip?.Dispose();
    }

    CompressionLevel _compression = CompressionLevel.Default;

    /// <summary>
    /// Compression level
    /// </summary>
    public CompressionLevel Compression
    {
        get => this._compression;
        set
        {
            foreach (ZipPackagePart? part in this.Parts.Values)
            {
                if (part.CompressionLevel == this._compression)
                {
                    part.CompressionLevel = value;
                }
            }

            this._compression = value;
        }
    }
}