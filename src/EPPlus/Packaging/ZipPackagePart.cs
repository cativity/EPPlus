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
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Packaging;

internal class ZipPackagePart : ZipPackagePartBase, IDisposable
{
    internal delegate void SaveHandlerDelegate(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName);

    internal ZipPackagePart(ZipPackage package, ZipEntry entry)
    {
        this.Package = package;
        this.Entry = entry;
        this.SaveHandler = null;
        this.Uri = new Uri(ZipPackage.GetUriKey(entry.FileName), UriKind.Relative);
    }
    internal ZipPackagePart(ZipPackage package, Uri partUri, string contentType, CompressionLevel compressionLevel)
    {
        this.Package = package;
        this.Uri = partUri;
        this.ContentType = contentType;
        this.CompressionLevel = compressionLevel;
    }
    internal ZipPackage Package { get; set; }
    internal ZipEntry Entry { get; set; }
    internal CompressionLevel CompressionLevel;
    Stream _stream = null;
    internal Stream Stream
    {
        get
        {
            return this._stream;
        }
        set
        {
            this._stream = value;
        }
    }
    internal override ZipPackageRelationship CreateRelationship(Uri targetUri, TargetMode targetMode, string relationshipType)
    {

        ZipPackageRelationship? rel = base.CreateRelationship(targetUri, targetMode, relationshipType);
        rel.SourceUri = this.Uri;
        return rel;
    }
    internal override ZipPackageRelationship CreateRelationship(string target, TargetMode targetMode, string relationshipType)
    {

        ZipPackageRelationship? rel = base.CreateRelationship(target, targetMode, relationshipType);
        rel.SourceUri = this.Uri;
        return rel;
    }

    internal Stream GetStream()
    {
        return this.GetStream(FileMode.OpenOrCreate, FileAccess.ReadWrite);
    }
    internal Stream GetStream(FileMode fileMode)
    {
        return this.GetStream(FileMode.Create, FileAccess.ReadWrite);
    }
    internal Stream GetStream(FileMode fileMode, FileAccess fileAccess)
    {
        if (this._stream == null || fileMode == FileMode.CreateNew || fileMode == FileMode.Create)
        {
            this._stream = RecyclableMemory.GetStream();
        }
        else
        {
            this._stream.Seek(0, SeekOrigin.Begin);
        }
        return this._stream;
    }

    string _contentType = "";
    public string ContentType
    {
        get
        {
            return this._contentType;
        }
        internal set
        {
            if (!string.IsNullOrEmpty(this._contentType))
            {
                if (this.Package._contentTypes.ContainsKey(ZipPackage.GetUriKey(this.Uri.OriginalString)))
                {
                    this.Package._contentTypes.Remove(ZipPackage.GetUriKey(this.Uri.OriginalString));
                    this.Package._contentTypes.Add(ZipPackage.GetUriKey(this.Uri.OriginalString), new ZipPackage.ContentType(value, false, this.Uri.OriginalString));
                }
            }

            this._contentType = value;
        }
    }
    public Uri Uri { get; private set; }
    public static Stream GetZipStream()
    {
        MemoryStream ms = RecyclableMemory.GetStream();
        ZipOutputStream os = new ZipOutputStream(ms);
        return os;
    }
    internal SaveHandlerDelegate SaveHandler
    {
        get;
        set;
    }
    internal void WriteZip(ZipOutputStream os)
    {
        byte[] b;
        if (this.SaveHandler == null)
        {
            b = ((MemoryStream)this.GetStream()).ToArray();
            if (b.Length == 0)   //Make sure the file isn't empty. DotNetZip streams does not seems to handle zero sized files.
            {
                return;
            }
            os.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)this.CompressionLevel;
            os.PutNextEntry(this.Uri.OriginalString);
            os.Write(b, 0, b.Length);
        }
        else
        {
            this.SaveHandler(os, (CompressionLevel)this.CompressionLevel, this.Uri.OriginalString);
        }

        if (this._rels.Count > 0)
        {
            string f = this.Uri.OriginalString;
            string? name = Path.GetFileName(f);
            this._rels.WriteZip(os, (string.Format("{0}_rels/{1}.rels", f.Substring(0, f.Length - name.Length), name)));
        }
        b = null;
    }


    public void Dispose()
    {
        this._stream.Close();
        this._stream.Dispose();
    }
}