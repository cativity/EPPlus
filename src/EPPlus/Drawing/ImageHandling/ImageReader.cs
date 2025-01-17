﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                   Change
 *************************************************************************************************
  12/03/2021         EPPlus Software AB       Added
 *************************************************************************************************/

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using OfficeOpenXml.Utils;
using System.Threading;
using OfficeOpenXml.Packaging.Ionic.Zlib;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Drawing
{
    internal class ImageReader
    {
        private const float M_TO_INCH = 39.3700787F;
        private const float CM_TO_INCH = M_TO_INCH * 0.01F;
        private const float MM_TO_INCH = CM_TO_INCH * 0.1F;
        private const float HUNDREDTH_TH_MM_TO_INCH = MM_TO_INCH * 0.01F;

        [SuppressMessage("ReSharper", "NotAccessedField.Global")]
        internal struct TifIfd
        {
            public short Tag;
            public short Type;
            public int Count;
            public int ValueOffset;
        }

        internal static ePictureType? GetPictureType(Stream stream, bool throwException)
        {
            ePictureType? pt;

            if (stream is MemoryStream ms)
            {
                pt = GetPictureTypeFromMs(ms);
            }
            else
            {
                Stream newMs = RecyclableMemory.GetStream();
                StreamUtil.CopyStream(stream, ref newMs);
                pt = GetPictureTypeFromMs((MemoryStream)newMs);
            }

            if (throwException && pt == null)
            {
                throw new InvalidOperationException("Cannot identify the image format of the stream.");
            }

            return pt;
        }
#if !NET35
        internal static async Task<ePictureType?> GetPictureTypeAsync(Stream stream)
        {
            ePictureType? pt;

            if (stream is MemoryStream ms)
            {
                pt = GetPictureTypeFromMs(ms);
            }
            else
            {
                Stream newMs = RecyclableMemory.GetStream();
                await StreamUtil.CopyStreamAsync(stream, newMs, new CancellationToken()).ConfigureAwait(false);
                pt = GetPictureTypeFromMs((MemoryStream)newMs);
            }

            if (pt == null)
            {
                throw new InvalidOperationException("Cannot identify the image format of the stream.");
            }

            return pt;
        }
#endif
        private static ePictureType? GetPictureTypeFromMs(MemoryStream ms)
        {
            BinaryReader? br = new BinaryReader(ms);

            if (IsJpg(br))
            {
                return ePictureType.Jpg;
            }

            if (IsBmp(br, out _))
            {
                return ePictureType.Bmp;
            }
            else if (IsGif(br))
            {
                return ePictureType.Gif;
            }
            else if (IsPng(br))
            {
                return ePictureType.Png;
            }
            else if (IsTif(br, out _, true))
            {
                return ePictureType.Tif;
            }
            else if (IsIco(br))
            {
                return ePictureType.Ico;
            }
            else if (IsWebP(br))
            {
                return ePictureType.WebP;
            }
            else if (IsEmf(br))
            {
                return ePictureType.Emf;
            }
            else if (IsWmf(br))
            {
                return ePictureType.Wmf;
            }
            else if (IsSvg(ms))
            {
                return ePictureType.Svg;
            }
            else if (IsGZip(br))
            {
                ms.Position = 0;
                _ = ExtractImage(ms.ToArray(), out ePictureType? pt);

                return pt;
            }

            return null;
        }

        private static bool IsGZip(BinaryReader br)
        {
            br.BaseStream.Position = 0;
            byte[]? sign = br.ReadBytes(2);

            return IsGZip(sign);
        }

        private static bool IsGZip(byte[] sign) => sign.Length >= 2 && sign[0] == 0x1F && sign[1] == 0x8B;

        internal static bool TryGetImageBounds(ePictureType pictureType,
                                               MemoryStream ms,
                                               ref double width,
                                               ref double height,
                                               out double horizontalResolution,
                                               out double verticalResolution)
        {
            width = 0;
            height = 0;
            horizontalResolution = verticalResolution = ExcelDrawing.STANDARD_DPI;

            try
            {
                _ = ms.Seek(0, SeekOrigin.Begin);

                if (pictureType == ePictureType.Bmp && IsBmp(ms, ref width, ref height, ref horizontalResolution, ref verticalResolution))
                {
                    return true;
                }

                if (pictureType == ePictureType.Jpg && IsJpg(ms, ref width, ref height, ref horizontalResolution, ref verticalResolution))
                {
                    return true;
                }

                if (pictureType == ePictureType.Gif && IsGif(ms, ref width, ref height))
                {
                    return true;
                }

                if (pictureType == ePictureType.Png && IsPng(ms, ref width, ref height, ref horizontalResolution, ref verticalResolution))
                {
                    return true;
                }

                if (pictureType == ePictureType.Emf && IsEmf(ms, ref width, ref height, ref horizontalResolution, ref verticalResolution))
                {
                    return true;
                }

                if (pictureType == ePictureType.Wmf && IsWmf(ms, ref width, ref height, ref horizontalResolution, ref verticalResolution))
                {
                    return true;
                }
                else if (pictureType == ePictureType.Svg && IsSvg(ms, ref width, ref height))
                {
                    return true;
                }
                else if (pictureType == ePictureType.Tif && IsTif(ms, ref width, ref height, ref horizontalResolution, ref verticalResolution))
                {
                    return true;
                }
                else if (pictureType == ePictureType.WebP && IsWebP(ms, ref width, ref height, ref horizontalResolution, ref verticalResolution))
                {
                    return true;
                }
                else if (pictureType == ePictureType.Ico && IsIcon(ms, ref width, ref height))
                {
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        internal static byte[] ExtractImage(byte[] img, out ePictureType? type)
        {
            if (IsGZip(img))
            {
                try
                {
                    MemoryStream? ms = RecyclableMemory.GetStream(img);
                    MemoryStream? msOut = RecyclableMemory.GetStream();
                    const int bufferSize = 4096;
                    byte[]? buffer = new byte[bufferSize];
                    using GZipStream? z = new GZipStream(ms, CompressionMode.Decompress);
                    int size;

                    do
                    {
                        size = z.Read(buffer, 0, bufferSize);

                        if (size > 0)
                        {
                            msOut.Write(buffer, 0, size);
                        }
                    } while (size == bufferSize);

                    msOut.Position = 0;
                    BinaryReader? br = new BinaryReader(msOut);

                    if (IsEmf(br))
                    {
                        type = ePictureType.Emf;
                    }
                    else if (IsWmf(br))
                    {
                        type = ePictureType.Wmf;
                    }
                    else
                    {
                        type = null;
                    }

                    msOut.Position = 0;

                    return msOut.ToArray();
                }
                catch
                {
                    type = null;

                    return img;
                }
            }

            type = null;

            return img;
        }

        private static bool IsJpg(MemoryStream ms, ref double width, ref double height, ref double horizontalResolution, ref double verticalResolution)
        {
            using BinaryReader? br = new BinaryReader(ms);

            if (IsJpg(br))
            {
                while (ms.Position < ms.Length)
                {
                    ushort id = GetUInt16BigEndian(br);
                    int length = (int)GetInt16BigEndian(br);

                    switch (id)
                    {
                        case 0xFFE0:
                            byte[]? identifier = br.ReadBytes(5); //JFIF\0
                            _ = br.ReadBytes(2);
                            byte unit = br.ReadByte();
                            float xDensity = (int)GetInt16BigEndian(br);
                            float yDensity = (int)GetInt16BigEndian(br);

                            if (unit == 1)
                            {
                                horizontalResolution = xDensity;
                                verticalResolution = yDensity;
                            }
                            else if (unit == 2)
                            {
                                horizontalResolution = xDensity * CM_TO_INCH;
                                verticalResolution = yDensity * CM_TO_INCH;
                            }

                            ms.Position += length - 14;

                            break;

                        case 0xFFE1:
                            long pos = ms.Position;
                            _ = br.ReadBytes(6); //EXIF\0\0 or //EXIF\FF\FF

                            double w = 0,
                                   h = 0;

                            _ = ReadTiffHeader(br, ref w, ref h, ref horizontalResolution, ref verticalResolution);
                            ms.Position = pos + length - 2;

                            break;

                        case 0xFFC0:
                        case 0xFFC1:
                        case 0xFFC2:
                            _ = br.ReadByte();
                            height = GetUInt16BigEndian(br);
                            width = GetUInt16BigEndian(br);
                            br.Close();

                            return true;

                        case 0xFFD9:
                            return height != 0 && width != 0;

                        default:
                            ms.Position += length - 2;

                            break;
                    }
                }
            }

            return false;
        }

        private static bool IsJpg(BinaryReader br)
        {
            br.BaseStream.Position = 0;
            byte[]? sign = br.ReadBytes(2); //FF D8 

            return sign.Length >= 2 && sign[0] == 0xFF && sign[1] == 0xD8;
        }

        private static bool IsGif(MemoryStream ms, ref double width, ref double height)
        {
            using BinaryReader? br = new BinaryReader(ms);

            if (IsGif(br))
            {
                width = br.ReadUInt16();
                height = br.ReadUInt16();
                br.Close();

                return true;
            }

            return false;
        }

        private static bool IsGif(BinaryReader br)
        {
            _ = br.BaseStream.Seek(0, SeekOrigin.Begin);
            byte[]? b = br.ReadBytes(6);

            return b[0] == 0x47 && b[1] == 0x49 && b[2] == 0x46; //byte 4-6 contains the version, but we don't check them here.
        }

        private static bool IsBmp(MemoryStream ms, ref double width, ref double height, ref double horizontalResolution, ref double verticalResolution)
        {
            using BinaryReader? br = new BinaryReader(ms);

            if (IsBmp(br, out string sign))
            {
                _ = br.ReadInt32();
                _ = br.ReadBytes(4);
                _ = br.ReadInt32();

                //Info Header
                _ = br.ReadInt32();
                width = br.ReadInt32();
                height = br.ReadInt32();

                if (sign == "BM")
                {
                    _ = br.ReadBytes(12);
                    horizontalResolution = br.ReadInt32() / M_TO_INCH;
                    verticalResolution = br.ReadInt32() / M_TO_INCH;
                }
                else
                {
                    horizontalResolution = verticalResolution = 1;
                }

                return true;
            }

            return false;
        }

        internal static bool IsBmp(BinaryReader br, out string sign)
        {
            try
            {
                _ = br.BaseStream.Seek(0, SeekOrigin.Begin);
                sign = Encoding.ASCII.GetString(br.ReadBytes(2)); //BM for a Windows bitmap

                return sign == "BM" || sign == "BA" || sign == "CI" || sign == "CP" || sign == "IC" || sign == "PT";
            }
            catch
            {
                sign = null;

                return false;
            }
        }

        #region Ico

        private static bool IsIcon(MemoryStream ms, ref double width, ref double height)
        {
            using BinaryReader? br = new BinaryReader(ms);

            if (IsIco(br))
            {
                _ = br.ReadInt16();
                width = br.ReadByte();

                if (width == 0)
                {
                    width = 256;
                }

                height = br.ReadByte();

                if (height == 0)
                {
                    height = 256;
                }

                //Icons will currently use the size from the icon and will not read the actual image size from the bmp or png. 

                //br.ReadBytes(6); //Ignore
                //var fileSize = br.ReadInt32();
                //if (fileSize > 0)
                //{
                //    var offset = br.ReadInt32();
                //    br.BaseStream.Position = offset;

                //    IsPng(br, ref width, ref height, offset+fileSize);
                //}
                br.Close();

                return true;
            }

            br.Close();

            return false;
        }

        internal static bool IsIco(BinaryReader br)
        {
            _ = br.BaseStream.Seek(0, SeekOrigin.Begin);
            short type0 = br.ReadInt16();
            short type1 = br.ReadInt16();

            return type0 == 0 && type1 == 1;
        }

        #endregion

        #region WebP

        private static bool IsWebP(MemoryStream ms, ref double width, ref double height, ref double horizontalResolution, ref double verticalResolution)
        {
            width = height = 0;
            horizontalResolution = verticalResolution = ExcelDrawing.STANDARD_DPI * (1 + (1 / 3)); //Excel seems to render webp at 1 1/3 size.

            using (BinaryReader? br = new BinaryReader(ms))
            {
                if (IsWebP(br))
                {
                    string? vp8 = Encoding.ASCII.GetString(br.ReadBytes(4));

                    switch (vp8)
                    {
                        case "VP8 ":
                            _ = br.ReadBytes(10);
                            short w = br.ReadInt16();
                            width = w & 0x3FFF;
                            //int hScale = w >> 14;
                            short h = br.ReadInt16();
                            height = h & 0x3FFF;
                            //hScale = h >> 14;

                            break;

                        case "VP8X":
                            _ = br.ReadBytes(8);
                            byte[] b = br.ReadBytes(6);
                            width = BitConverter.ToInt32(new byte[] { b[0], b[1], b[2], 0 }, 0) + 1;
                            height = BitConverter.ToInt32(new byte[] { b[3], b[4], b[5], 0 }, 0) + 1;

                            break;

                        case "VP8L":
                            _ = br.ReadBytes(5);
                            b = br.ReadBytes(4);
                            width = (b[0] | ((b[1] & 0x3F) << 8)) + 1;
                            height = ((b[1] >> 6) | (b[2] << 2) | ((b[3] & 0x0F) << 10)) + 1;

                            break;
                    }
                }
            }

            return width != 0 && height != 0;
        }

        internal static bool IsWebP(BinaryReader br)
        {
            try
            {
                _ = br.BaseStream.Seek(0, SeekOrigin.Begin);
                string? riff = Encoding.ASCII.GetString(br.ReadBytes(4));
                _ = GetInt32BigEndian(br);
                string? webP = Encoding.ASCII.GetString(br.ReadBytes(4));

                return riff == "RIFF" && webP == "WEBP";
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Tiff

        private static bool IsTif(MemoryStream ms, ref double width, ref double height, ref double horizontalResolution, ref double verticalResolution)
        {
            using BinaryReader? br = new BinaryReader(ms);

            return ReadTiffHeader(br, ref width, ref height, ref horizontalResolution, ref verticalResolution);
        }

        private static bool ReadTiffHeader(BinaryReader br, ref double width, ref double height, ref double horizontalResolution, ref double verticalResolution)
        {
            Stream? ms = br.BaseStream;
            long pos = ms.Position;

            if (IsTif(br, out bool isBigEndian, false))
            {
                int offset = GetTifInt32(br, isBigEndian);
                ms.Position = pos + offset;
                short numberOfIdf = GetTifInt16(br, isBigEndian);
                List<TifIfd>? ifds = new List<TifIfd>();

                for (int i = 0; i < numberOfIdf; i++)
                {
                    TifIfd ifd = new TifIfd()
                    {
                        Tag = GetTifInt16(br, isBigEndian), Type = GetTifInt16(br, isBigEndian), Count = GetTifInt32(br, isBigEndian),
                    };

                    if (ifd.Type == 1 || ifd.Type == 2 || ifd.Type == 6 || ifd.Type == 7)
                    {
                        ifd.ValueOffset = br.ReadByte();
                        _ = br.ReadBytes(3);
                    }
                    else if (ifd.Type == 3 || ifd.Type == 8)
                    {
                        ifd.ValueOffset = GetTifInt16(br, isBigEndian);
                        _ = br.ReadBytes(2);
                    }
                    else
                    {
                        ifd.ValueOffset = GetTifInt32(br, isBigEndian);
                    }

                    ifds.Add(ifd);
                }

                int resolutionUnit = 2;

                foreach (TifIfd ifd in ifds)
                {
                    switch (ifd.Tag)
                    {
                        case 0x100:
                            width = ifd.ValueOffset;

                            break;

                        case 0x101:
                            height = ifd.ValueOffset;

                            break;

                        case 0x11A:
                            ms.Position = ifd.ValueOffset + pos;
                            int l1 = GetTifInt32(br, isBigEndian);
                            int l2 = GetTifInt32(br, isBigEndian);
                            horizontalResolution = l1 / l2;

                            break;

                        case 0x11B:
                            ms.Position = ifd.ValueOffset + pos;
                            l1 = GetTifInt32(br, isBigEndian);
                            l2 = GetTifInt32(br, isBigEndian);
                            verticalResolution = l1 / l2;

                            break;

                        case 0x128:
                            resolutionUnit = ifd.ValueOffset;

                            break;
                    }
                }

                if (resolutionUnit == 1)
                {
                    horizontalResolution *= CM_TO_INCH;
                    verticalResolution *= CM_TO_INCH;
                }
            }

            return width != 0 && height != 0;
        }

        private static bool IsTif(BinaryReader br, out bool isBigEndian, bool resetPos)
        {
            try
            {
                if (resetPos)
                {
                    br.BaseStream.Position = 0;
                }

                byte[]? b = br.ReadBytes(2);
                isBigEndian = Encoding.ASCII.GetString(b) == "MM";
                short identifier = GetTifInt16(br, isBigEndian);

                if (identifier == 42)
                {
                    return true;
                }
            }
            catch
            {
                isBigEndian = false;

                return false;
            }

            return false;
        }

        private static short GetTifInt16(BinaryReader br, bool isBigEndian)
        {
            if (isBigEndian)
            {
                return GetInt16BigEndian(br);
            }
            else
            {
                return br.ReadInt16();
            }
        }

        private static int GetTifInt32(BinaryReader br, bool isBigEndian)
        {
            if (isBigEndian)
            {
                return GetInt32BigEndian(br);
            }
            else
            {
                return br.ReadInt32();
            }
        }

        #endregion

        #region Emf

        private static bool IsEmf(MemoryStream ms, ref double width, ref double height, ref double horizontalResolution, ref double verticalResolution)
        {
            using BinaryReader? br = new BinaryReader(ms);

            if (IsEmf(br))
            {
                _ = br.ReadInt32();
                int[]? bounds = new int[4];
                bounds[0] = br.ReadInt32();
                bounds[1] = br.ReadInt32();
                bounds[2] = br.ReadInt32();
                bounds[3] = br.ReadInt32();
                int[]? frame = new int[4];
                frame[0] = br.ReadInt32();
                frame[1] = br.ReadInt32();
                frame[2] = br.ReadInt32();
                frame[3] = br.ReadInt32();

                byte[]? signatureBytes = br.ReadBytes(4);
                string? signature = Encoding.ASCII.GetString(signatureBytes);

                if (signature.Trim() == "EMF")
                {
                    _ = br.ReadUInt32();
                    _ = br.ReadUInt32();
                    _ = br.ReadUInt32();
                    _ = br.ReadUInt16();
                    _ = br.ReadUInt16();

                    _ = br.ReadUInt32();
                    _ = br.ReadUInt32();
                    _ = br.ReadUInt32();
                    uint[]? device = new uint[2];
                    device[0] = br.ReadUInt32();
                    device[1] = br.ReadUInt32();

                    uint[]? mm = new uint[2];
                    mm[0] = br.ReadUInt32();
                    mm[1] = br.ReadUInt32();

                    //Extension 1
                    _ = br.ReadUInt32();
                    _ = br.ReadUInt32();
                    _ = br.ReadUInt32();

                    //Extension 2
                    _ = br.ReadUInt32();
                    _ = br.ReadUInt32();

                    _ = br.ReadInt32();
                    _ = br.ReadInt32();

                    width = bounds[2] - bounds[0] + 1;
                    height = bounds[3] - bounds[1] + 1;

                    horizontalResolution = width / ((frame[2] - frame[0]) * HUNDREDTH_TH_MM_TO_INCH * ExcelDrawing.STANDARD_DPI) * ExcelDrawing.STANDARD_DPI;
                    verticalResolution = height / ((frame[3] - frame[1]) * HUNDREDTH_TH_MM_TO_INCH * ExcelDrawing.STANDARD_DPI) * ExcelDrawing.STANDARD_DPI;

                    return true;
                }
            }

            return false;
        }

        private static bool IsEmf(BinaryReader br)
        {
            br.BaseStream.Position = 0;
            int type = br.ReadInt32();

            return type == 1;
        }

        #endregion

        #region Wmf

        private const double PIXELS_PER_TWIPS = 1D / 15D;
        private const double DEFAULT_TWIPS = 1440D;

        private static bool IsWmf(MemoryStream ms, ref double width, ref double height, ref double horizontalResolution, ref double verticalResolution)
        {
            using BinaryReader? br = new BinaryReader(ms);

            if (IsWmf(br))
            {
                _ = br.ReadInt16();
                ushort[]? bounds = new ushort[4];
                bounds[0] = br.ReadUInt16();
                bounds[1] = br.ReadUInt16();
                bounds[2] = br.ReadUInt16();
                bounds[3] = br.ReadUInt16();

                short inch = br.ReadInt16();
                width = bounds[2] - bounds[0];
                height = bounds[3] - bounds[1];

                if (inch != 0)
                {
                    width *= DEFAULT_TWIPS / inch * PIXELS_PER_TWIPS;
                    height *= DEFAULT_TWIPS / inch * PIXELS_PER_TWIPS;
                }

                return width != 0 && height != 0;
            }

            return false;
        }

        private static bool IsWmf(BinaryReader br)
        {
            br.BaseStream.Position = 0;
            uint key = br.ReadUInt32();

            return key == 0x9AC6CDD7;
        }

        #endregion

        #region Png

        private static bool IsPng(MemoryStream ms, ref double width, ref double height, ref double horizontalResolution, ref double verticalResolution)
        {
            using BinaryReader? br = new BinaryReader(ms);

            return IsPng(br, ref width, ref height, ref horizontalResolution, ref verticalResolution);
        }

        private static bool IsPng(BinaryReader br,
                                  ref double width,
                                  ref double height,
                                  ref double horizontalResolution,
                                  ref double verticalResolution,
                                  long fileEndPosition = long.MinValue)
        {
            if (IsPng(br))
            {
                if (fileEndPosition == long.MinValue)
                {
                    fileEndPosition = br.BaseStream.Length;
                }

                while (br.BaseStream.Position < fileEndPosition)
                {
                    string? chunkType = ReadPngChunkHeader(br, out int length);

                    switch (chunkType)
                    {
                        case "IHDR":
                            width = GetInt32BigEndian(br);
                            height = GetInt32BigEndian(br);
                            _ = br.ReadBytes(5); //Ignored bytes, Depth compression etc.

                            break;

                        case "pHYs":
                            horizontalResolution = GetInt32BigEndian(br);
                            verticalResolution = GetInt32BigEndian(br);
                            byte unitSpecifier = br.ReadByte();

                            if (unitSpecifier == 1)
                            {
                                horizontalResolution /= M_TO_INCH;
                                verticalResolution /= M_TO_INCH;
                            }

                            br.Close();

                            return true;

                        case "IEND":
                            br.Close();

                            return width != 0 && height != 0;

                        default:
                            _ = br.ReadBytes(length);

                            break;
                    }

                    _ = br.ReadInt32();
                }
            }

            br.Close();

            return width != 0 && height != 0;
        }

        private static bool IsPng(BinaryReader br)
        {
            br.BaseStream.Position = 0;
            byte[]? signature = br.ReadBytes(8);

            return signature.SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 });
        }

        private static string ReadPngChunkHeader(BinaryReader br, out int length)
        {
            length = GetInt32BigEndian(br);
            byte[]? b = br.ReadBytes(4);
            string? type = Encoding.ASCII.GetString(b);

            return type;
        }

        #endregion

        #region Svg

        private static bool IsSvg(MemoryStream ms, ref double width, ref double height)
        {
            try
            {
                using XmlTextReader? reader = new XmlTextReader(ms);

                while (reader.Read())
                {
                    if (reader.LocalName == "svg" && reader.NodeType == XmlNodeType.Element)
                    {
                        string? w = reader.GetAttribute("width");
                        string? h = reader.GetAttribute("height");
                        string? vb = reader.GetAttribute("viewBox");
                        reader.Close();

                        if (w == null || h == null)
                        {
                            if (vb == null)
                            {
                                return false;
                            }

                            string[]? bounds = vb.Split(new char[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);

                            if (bounds.Length < 4)
                            {
                                return false;
                            }

                            if (string.IsNullOrEmpty(w))
                            {
                                w = bounds[2];
                            }

                            if (string.IsNullOrEmpty(h))
                            {
                                h = bounds[3];
                            }
                        }

                        width = GetSvgUnit(w);

                        if (double.IsNaN(width))
                        {
                            return false;
                        }

                        height = GetSvgUnit(h);

                        if (double.IsNaN(height))
                        {
                            return false;
                        }

                        return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsSvg(MemoryStream ms)
        {
            try
            {
                ms.Position = 0;
                XmlTextReader? reader = new XmlTextReader(ms);

                while (reader.Read())
                {
                    if (reader.LocalName == "svg" && reader.NodeType == XmlNodeType.Element)
                    {
                        return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private static double GetSvgUnit(string v)
        {
            double factor = 1D;

            if (v.EndsWith("px", StringComparison.OrdinalIgnoreCase))
            {
                v = v.Substring(0, v.Length - 2);
            }
            else if (v.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            {
                factor = 1.25;
                v = v.Substring(0, v.Length - 2);
            }
            else if (v.EndsWith("pc", StringComparison.OrdinalIgnoreCase))
            {
                factor = 15;
                v = v.Substring(0, v.Length - 2);
            }
            else if (v.EndsWith("mm", StringComparison.OrdinalIgnoreCase))
            {
                factor = 3.543307;
                v = v.Substring(0, v.Length - 2);
            }
            else if (v.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
            {
                factor = 35.43307;
                v = v.Substring(0, v.Length - 2);
            }
            else if (v.EndsWith("in", StringComparison.OrdinalIgnoreCase))
            {
                factor = 90;
                v = v.Substring(0, v.Length - 2);
            }

            if (double.TryParse(v, out double value))
            {
                return value * factor;
            }

            return double.NaN;
        }

        #endregion

        private static ushort GetUInt16BigEndian(BinaryReader br)
        {
            byte[]? b = br.ReadBytes(2);

            return BitConverter.ToUInt16(new byte[] { b[1], b[0] }, 0);
        }

        private static short GetInt16BigEndian(BinaryReader br)
        {
            byte[]? b = br.ReadBytes(2);

            return BitConverter.ToInt16(new byte[] { b[1], b[0] }, 0);
        }

        private static int GetInt32BigEndian(BinaryReader br)
        {
            byte[]? b = br.ReadBytes(4);

            return BitConverter.ToInt32(new byte[] { b[3], b[2], b[1], b[0] }, 0);
        }
    }
}