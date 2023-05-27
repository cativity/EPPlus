﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/

using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;

/// <summary>
/// Loads serialized font metrics
/// </summary>
internal static class GenericFontMetricsLoader
{
    /// <summary>
    /// Loads all serialized font metrics from the resources/SerializedFonts.zip archive
    /// </summary>
    internal static Dictionary<uint, SerializedFontMetrics> LoadFontMetrics()
    {
        Dictionary<uint, SerializedFontMetrics>? fonts = new Dictionary<uint, SerializedFontMetrics>();
        Assembly? assembly = Assembly.GetExecutingAssembly();

        using (Stream? stream = assembly.GetManifestResourceStream("OfficeOpenXml.resources.TextMetrics.zip"))
        {
            ZipInputStream? zipStream = new ZipInputStream(stream);
            ZipEntry entry;

            while ((entry = zipStream.GetNextEntry()) != null)
            {
                if (!entry.IsDirectory && Path.GetExtension(entry.FileName) == ".fmtr")
                {
                    byte[]? bytes = new byte[entry.UncompressedSize];
                    _ = zipStream.Read(bytes, 0, (int)entry.UncompressedSize);
                    using MemoryStream? ms = RecyclableMemory.GetStream(bytes);
                    SerializedFontMetrics? fnt = GenericFontMetricsSerializer.Deserialize(ms);
                    fonts.Add(fnt.GetKey(), fnt);
                }
            }
        }

        return fonts;
    }
}