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
using System.Runtime.InteropServices;
using comTypes = System.Runtime.InteropServices.ComTypes;
using System.IO;
using System.Security;

namespace OfficeOpenXml.Utils.CompundDocument;

internal class CompoundDocument
{
    internal class StoragePart
    {
        public StoragePart()
        {
        }

        internal Dictionary<string, StoragePart> SubStorage = new Dictionary<string, StoragePart>();
        internal Dictionary<string, byte[]> DataStreams = new Dictionary<string, byte[]>();
    }

    /// <summary>
    /// The root storage part of the compound document.
    /// </summary>
    internal StoragePart Storage;

    /// <summary>
    /// Directories in the order they are saved.
    /// </summary>
    internal List<CompoundDocumentItem> Directories { get; private set; }

    internal CompoundDocument() => this.Storage = new StoragePart();

    internal CompoundDocument(MemoryStream ms) => this.Read(ms);

    internal CompoundDocument(FileInfo fi) => this.Read(fi);

    internal static bool IsCompoundDocument(FileInfo fi) => CompoundDocumentFile.IsCompoundDocument(fi);

    internal static bool IsCompoundDocument(MemoryStream ms) => CompoundDocumentFile.IsCompoundDocument(ms);

    internal CompoundDocument(byte[] doc) => this.Read(doc);

    internal void Read(FileInfo fi)
    {
        byte[]? b = File.ReadAllBytes(fi.FullName);
        this.Read(b);
    }

    internal void Read(byte[] doc)
    {
        using MemoryStream? ms = RecyclableMemory.GetStream(doc);
        this.Read(ms);
    }

    internal void Read(MemoryStream ms)
    {
        using CompoundDocumentFile? doc = new CompoundDocumentFile(ms);
        this.Storage = new StoragePart();
        GetStorageAndStreams(this.Storage, doc.RootItem);
        this.Directories = doc.Directories;
    }

    private static void GetStorageAndStreams(StoragePart storage, CompoundDocumentItem parent)
    {
        foreach (CompoundDocumentItem? item in parent.Children)
        {
            if (item.ObjectType == 1) //Substorage
            {
                StoragePart? part = new StoragePart();
                storage.SubStorage.Add(item.Name, part);
                GetStorageAndStreams(part, item);
            }
            else if (item.ObjectType == 2) //Stream
            {
                storage.DataStreams.Add(item.Name, item.Stream);
            }
        }
    }

    internal void Save(MemoryStream ms)
    {
        CompoundDocumentFile? doc = new CompoundDocumentFile();
        WriteStorageAndStreams(this.Storage, doc.RootItem);
        this.Directories = doc.Directories;
        doc.Write(ms);
    }

    private static void WriteStorageAndStreams(StoragePart storage, CompoundDocumentItem parent)
    {
        foreach (KeyValuePair<string, StoragePart> item in storage.SubStorage)
        {
            CompoundDocumentItem? c = new CompoundDocumentItem() { Name = item.Key, ObjectType = 1, Stream = null, StreamSize = 0, Parent = parent };
            parent.Children.Add(c);
            WriteStorageAndStreams(item.Value, c);
        }

        foreach (KeyValuePair<string, byte[]> item in storage.DataStreams)
        {
            CompoundDocumentItem? c = new CompoundDocumentItem()
            {
                Name = item.Key, ObjectType = 2, Stream = item.Value, StreamSize = item.Value == null ? 0 : item.Value.Length, Parent = parent
            };

            parent.Children.Add(c);
        }
    }
}