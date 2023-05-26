/*************************************************************************************************
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
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;

internal static class GenericFontMetricsSerializer
{
    public static readonly Encoding FileEncoding = Encoding.UTF8;

    public static SerializedFontMetrics Deserialize(Stream stream)
    {
        using BinaryReader? reader = new BinaryReader(stream, FileEncoding);
        SerializedFontMetrics? metrics = new SerializedFontMetrics();
        metrics.Version = reader.ReadUInt16();
        metrics.Family = (FontMetricsFamilies)reader.ReadUInt16();
        metrics.SubFamily = (FontSubFamilies)reader.ReadUInt16();
        metrics.LineHeight1em = reader.ReadSingle();
        metrics.DefaultWidthClass = (FontMetricsClass)reader.ReadByte();
        ushort nClassWidths = reader.ReadUInt16();
        if (nClassWidths == 0)
        {
            return metrics;
        }
        for (int x = 0; x < nClassWidths; x++)
        {
            FontMetricsClass cls = (FontMetricsClass)reader.ReadByte();
            float width = reader.ReadSingle();
            metrics.ClassWidths[cls] = width;
        }
        ushort nClasses = reader.ReadUInt16();
        for (int x = 0; x < nClasses; x++)
        {
            FontMetricsClass cls = (FontMetricsClass)reader.ReadByte();
            ushort nRanges = reader.ReadUInt16();
            for (int rngIx = 0; rngIx < nRanges; rngIx++)
            {
                ushort start = reader.ReadUInt16();
                ushort end = reader.ReadUInt16();
                for (ushort c = start; c <= end; c++)
                {
                    metrics.CharMetrics[Convert.ToChar(c)] = cls;
                }
            }
            ushort nCharactersInClass = reader.ReadUInt16();
            if (nCharactersInClass == 0)
            {
                continue;
            }

            for (int y = 0; y < nCharactersInClass; y++)
            {
                ushort cCode = reader.ReadUInt16();
                char c = Convert.ToChar(cCode);
                metrics.CharMetrics[c] = cls;
            }
        }
        return metrics;
    }
}