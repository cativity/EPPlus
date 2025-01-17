﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/05/2022         EPPlus Software AB       EPPlus 6.1
 *************************************************************************************************/

using OfficeOpenXml.Utils.CompundDocument;
using OfficeOpenXml.VBA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Vba.ContentHash;

internal class FormsNormalizedDataHashInputProvider : ContentHashInputProvider
{
    public FormsNormalizedDataHashInputProvider(ExcelVbaProject project)
        : base(project)
    {
    }

    protected override void CreateHashInputInternal(MemoryStream ms)
    {
        //MS-OVBA 2.4.2.2
        BinaryWriter bw = new BinaryWriter(ms);
        this.FormsNormaizedData(bw);
    }

    private void FormsNormaizedData(BinaryWriter bw)
    {
        ExcelVbaProject? p = this.Project;
        IList<string>? designers = GetDesigners(p);
        List<SortItem>? list = new List<SortItem>();

        foreach (string? designer in designers)
        {
            AppendDesignerStreams(p, list, designer);
        }

        WriteDesignerStreams(bw, p, list);
    }

    private static void WriteDesignerStreams(BinaryWriter bw, ExcelVbaProject p, List<SortItem> list)
    {
        HashSet<string>? hs = new HashSet<string>(list.Where(x => x.IsStream).Select(x => x.Name));

        foreach (CompoundDocumentItem? dir in p.Document.Directories)
        {
            if (hs.Contains(dir.FullName) && dir.StreamSize > 0)
            {
                WriteStreamData(bw, dir.Stream);
            }
        }
    }

    internal static void NormalizeDesigner(ExcelVbaProject p, BinaryWriter bw, string designer)
    {
        List<SortItem>? list = new List<SortItem>();
        AppendDesignerStreams(p, list, designer);
        WriteDesignerStreams(bw, p, list);
    }

    private static void AppendDesignerStreams(ExcelVbaProject p, List<SortItem> list, string designer)
    {
        CompoundDocument.StoragePart? storage = p.Document.Storage.SubStorage[designer];
        NormalizeStorage(storage, list, p.Document.Directories[0].Name + "/" + designer);
    }

    private static void NormalizeStorage(CompoundDocument.StoragePart storage, List<SortItem> list, string parentName)
    {
        IList<SortItem>? children = GetSortedChildren(storage);

        foreach (SortItem? child in children)
        {
            string? newChildName = parentName + "/" + child.Name;

            if (child.IsStream == false)
            {
                NormalizeStorage(storage.SubStorage[child.Name], list, newChildName);
            }

            child.Name = newChildName;
            list.Add(child);
        }
    }

    private static void WriteStreamData(BinaryWriter bw, byte[] b)
    {
        int streamLength;

        if (b != null)
        {
            bw.Write(b);
            streamLength = b.Length;
        }
        else
        {
            streamLength = 0;
        }

        int zeros = 1023 - (streamLength % 1023);

        for (int i = 0; i < zeros; i++)
        {
            bw.Write((byte)0);
        }
    }

    private class SortItem
    {
        public SortItem(string name, bool isStream)
        {
            this.Name = name;
            this.IsStream = isStream;
        }

        public string Name { get; set; }

        public bool IsStream { get; set; }
    }

    private static IList<SortItem> GetSortedChildren(CompoundDocument.StoragePart storage)
    {
        List<SortItem>? list = new List<SortItem>();
        list.AddRange(storage.DataStreams.Keys.Select(x => new SortItem(x, true)));
        list.AddRange(storage.SubStorage.Keys.Select(x => new SortItem(x, false)));

        return list;
    }

    private static IList<string> GetDesigners(ExcelVbaProject p)
    {
        IEnumerable<string>? designerModules = p.Modules.Where(x => x.Type == eModuleType.Designer).Select(x => x.streamName);

        return designerModules.ToList();
    }

    private static void NormalizeDesignerStorage(ExcelVBAModule designerModule, BinaryWriter bw)
    {
        //_ = new BufferedStream(bw.BaseStream, 1023);

        //var ds = p.Document.Storage.SubStorage[designerModule.streamName];
    }
}