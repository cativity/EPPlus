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
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils;

internal static class VBACompression
{
    #region Compression

    /// <summary>
    /// Compression using a run length encoding algorithm.
    /// See MS-OVBA Section 2.4
    /// </summary>
    /// <param name="part">Byte array to decompress</param>
    /// <returns></returns>
    internal static byte[] CompressPart(byte[] part)
    {
        using MemoryStream? ms = RecyclableMemory.GetStream(4096);
        BinaryWriter br = new BinaryWriter(ms);
        br.Write((byte)1);

        int compStart = 1;
        int compEnd = 4098;
        int decompStart = 0;
        int decompEnd = part.Length < 4096 ? part.Length : 4096;

        while (decompStart < decompEnd && compStart < compEnd)
        {
            byte[] chunk = CompressChunk(part, ref decompStart);

            if (chunk == null || chunk.Length == 0)
            {
                //header = 4096 | 0x600; //B=011 A=0
            }
            else
            {
                ushort header = (ushort)((chunk.Length - 1) & 0xFFF);
                header |= 0xB000; //B=011 A=1
                br.Write(header);
                br.Write(chunk);
            }

            decompEnd = part.Length < decompStart + 4096 ? part.Length : decompStart + 4096;
        }

        br.Flush();

        return ms.ToArray();
    }

    private static byte[] CompressChunk(byte[] buffer, ref int startPos)
    {
        byte[]? comprBuffer = new byte[4096];
        int flagPos = 0;
        int cPos = 1;
        int dPos = startPos;
        int dEnd = startPos + 4096 < buffer.Length ? startPos + 4096 : buffer.Length;

        while (dPos < dEnd)
        {
            byte tokenFlags = 0;

            for (int i = 0; i < 8; i++)
            {
                if (dPos - startPos > 0)
                {
                    int bestCandidate = -1;
                    int bestLength = 0;
                    int candidate = dPos - 1;
                    int bitCount = GetLengthBits(dPos - startPos);
                    int bits = 16 - bitCount;
                    ushort lengthMask = (ushort)(0xFFFF >> bits);

                    while (candidate >= startPos)
                    {
                        if (buffer[candidate] == buffer[dPos])
                        {
                            int length = 1;

                            while (buffer.Length > dPos + length
                                   && buffer[candidate + length] == buffer[dPos + length]
                                   && length < lengthMask
                                   && dPos + length < dEnd)
                            {
                                length++;
                            }

                            if (length > bestLength)
                            {
                                bestCandidate = candidate;
                                bestLength = length;

                                if (bestLength == lengthMask)
                                {
                                    break;
                                }
                            }
                        }

                        candidate--;
                    }

                    if (bestLength >= 3) //Copy token
                    {
                        tokenFlags |= (byte)(1 << i);

                        ushort token = (ushort)(((ushort)(dPos - (bestCandidate + 1)) << bitCount) | (ushort)(bestLength - 3));
                        Array.Copy(BitConverter.GetBytes(token), 0, comprBuffer, cPos, 2);
                        dPos += bestLength;
                        cPos += 2;

                        //SetCopy Token                        
                    }
                    else
                    {
                        comprBuffer[cPos++] = buffer[dPos++];
                    }
                }

                else
                {
                    comprBuffer[cPos++] = buffer[dPos++];
                }

                if (dPos >= dEnd)
                {
                    break;
                }
            }

            comprBuffer[flagPos] = tokenFlags;
            flagPos = cPos++;
        }

        byte[]? ret = new byte[cPos - 1];
        Array.Copy(comprBuffer, ret, ret.Length);
        startPos = dEnd;

        return ret;
    }

    internal static byte[] DecompressPart(byte[] part) => DecompressPart(part, 0);

    /// <summary>
    /// Decompression using a run length encoding algorithm.
    /// See MS-OVBA Section 2.4
    /// </summary>
    /// <param name="part">Byte array to decompress</param>
    /// <param name="startPos"></param>
    /// <returns></returns>
    internal static byte[] DecompressPart(byte[] part, int startPos)
    {
        if (part[startPos] != 1)
        {
            return null;
        }

        using MemoryStream? ms = RecyclableMemory.GetStream(4096);
        int compressPos = startPos + 1;

        while (compressPos < part.Length - 1)
        {
            DecompressChunk(ms, part, ref compressPos);
        }

        return ms.ToArray();
    }

    private static void DecompressChunk(MemoryStream ms, byte[] compBuffer, ref int pos)
    {
        ushort header = BitConverter.ToUInt16(compBuffer, pos);
        int decomprPos = 0;
        byte[] buffer = new byte[4198]; //Add an extra 100 byte. Some workbooks have overflowing worksheets.
        int size = (int)(header & 0xFFF) + 3;
        int endPos = pos + size;
        int b = (int)(header & 0x8000) >> 15;
        pos += 2;

        if (b == 1) //Compressed chunk
        {
            while (pos < compBuffer.Length && pos < endPos)
            {
                //Decompress token
                byte token = compBuffer[pos++];

                if (pos >= endPos)
                {
                    break;
                }

                for (int i = 0; i < 8; i++)
                {
                    //Literal token
                    if ((token & (1 << i)) == 0)
                    {
                        ms.WriteByte(compBuffer[pos]);
                        buffer[decomprPos++] = compBuffer[pos++];
                    }
                    else //copy token
                    {
                        ushort t = BitConverter.ToUInt16(compBuffer, pos);
                        int bitCount = GetLengthBits(decomprPos);
                        int bits = 16 - bitCount;
                        ushort lengthMask = (ushort)(0xFFFF >> bits);
                        ushort offsetMask = (ushort)~lengthMask;
                        int length = (lengthMask & t) + 3;
                        int offset = (offsetMask & t) >> bitCount;
                        int source = decomprPos - offset - 1;

                        if (decomprPos + length >= buffer.Length)
                        {
                            // Be lenient on decompression, so extend our decompression
                            // buffer. Excel generated VBA projects do encounter this issue.
                            // One would think (not surprisingly that the VBA project spec)
                            // over emphasizes the size restrictions of a DecompressionChunk.
                            byte[]? largerBuffer = new byte[buffer.Length + 4098];
                            Array.Copy(buffer, largerBuffer, decomprPos);
                            buffer = largerBuffer;
                        }

                        // Even though we've written to the MemoryStream,
                        // We still should decompress the token into this buffer
                        // in case a later token needs to use the bytes we're
                        // about to decompress.
                        for (int c = 0; c < length; c++)
                        {
                            ms.WriteByte(buffer[source]); //Must copy byte-wise because copytokens can overlap compressed buffer.
                            buffer[decomprPos++] = buffer[source++];
                        }

                        pos += 2;
                    }

                    if (pos >= endPos)
                    {
                        break;
                    }
                }
            }

            return;
        }
        else //Raw chunk
        {
            ms.Write(compBuffer, pos, size);
            pos += size;

            return;
        }
    }

    private static int GetLengthBits(int decompPos)
    {
        if (decompPos <= 16)
        {
            return 12;
        }
        else if (decompPos <= 32)
        {
            return 11;
        }
        else if (decompPos <= 64)
        {
            return 10;
        }
        else if (decompPos <= 128)
        {
            return 9;
        }
        else if (decompPos <= 256)
        {
            return 8;
        }
        else if (decompPos <= 512)
        {
            return 7;
        }
        else if (decompPos <= 1024)
        {
            return 6;
        }
        else if (decompPos <= 2048)
        {
            return 5;
        }
        else if (decompPos <= 4096)
        {
            return 4;
        }
        else
        {
            //We should never end up here, but if so this is the formula to calculate the bits...
            return 12 - (int)Math.Truncate(Math.Log((decompPos - 1) >> 4, 2) + 1);
        }
    }

    #endregion
}