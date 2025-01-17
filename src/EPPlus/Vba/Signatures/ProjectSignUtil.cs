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

using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA.ContentHash;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace OfficeOpenXml.VBA.Signatures;

internal static class ProjectSignUtil
{
    internal static ContentInfo SignProject(ExcelVbaProject proj, EPPlusVbaSignature signature, EPPlusSignatureContext ctx)
    {
        X509Certificate2? certificate = signature.Certificate;

        if (!certificate.HasPrivateKey)
        {
            //throw (new InvalidOperationException("The certificate doesn't have a private key"));
            signature.Certificate = null;

            return null;
        }

        byte[]? hash = VbaSignHashAlgorithmUtil.GetContentHash(proj, ctx);

        ContentInfo contentInfo;

        using (MemoryStream? ms = RecyclableMemory.GetStream())
        {
            contentInfo = CreateContentInfo(hash, ms, ctx);
        }

        return contentInfo;
    }

    private static ContentInfo CreateContentInfo(byte[] hash, MemoryStream ms, EPPlusSignatureContext ctx)
    {
        // [MS-OSHARED] 2.3.2.4.3.1 SpcIndirectDataContent
        BinaryWriter bw = new BinaryWriter(ms);

        byte[]? hashAlgorithmBytes = ctx.GetHashAlgorithmBytes();
        byte[]? hashContentBytes = GetHashContent(ctx, hash);

        bw.Write((byte)0x30); //Constructed Type 

        if (ctx.SignatureType == ExcelVbaSignatureType.Legacy)
        {
            bw.Write((byte)(hashAlgorithmBytes.Length + hashContentBytes.Length + 0x18)); //Total length
        }
        else
        {
            byte length = (byte)(hashAlgorithmBytes.Length + hashContentBytes.Length + 0x24);
            WriteSequenceLength(bw, length);
        }

        bw.Write((byte)0x30); //Constructed Type 
        bw.Write((byte)0x0E); //Length SpcIndirectDataContent

        byte[]? spcIndirectDataContentOidBytes = ctx.GetIndirectDataContentOidBytes();
        WriteOid(bw, spcIndirectDataContentOidBytes);

        if (ctx.SignatureType == ExcelVbaSignatureType.Legacy)
        {
            bw.Write((byte)0x04); //Octet String Tag Identifier
            bw.Write((byte)0x00); //Zero length
        }
        else
        {
            // SigFormatDescriptorV1
            bw.Write((byte)0x04);
            bw.Write((byte)0x0C); // Size of octstring
            bw.Write(12); // size of record
            bw.Write(1); // version
            bw.Write(1); // format
        }

        bw.Write((byte)0x30); //Constructed Type (DigestInfo)
        bw.Write((byte)0x20); //Length DigestInfo
        bw.Write((byte)0x30); //Constructed Type (Algorithm)
        bw.Write((byte)(hashAlgorithmBytes.Length + 7)); //length AlgorithmIdentifier

        WriteOid(bw, hashAlgorithmBytes); //Hash Algorithem

        bw.Write((byte)0x05); //Null type identifier
        bw.Write((byte)0x00); //Null length
        bw.Write(hashContentBytes);

        ContentInfo contentInfo = new(new Oid("1.3.6.1.4.1.311.2.1.4"), ms.ToArray());

        return contentInfo;
    }

    private static void WriteSequenceLength(BinaryWriter bw, int length)
    {
        int num = length;

        if (num >= 0x80)
        {
            int bytes = GetByteSize(length);
            length += bytes;
            int b2 = GetByteSize(length);

            if (bytes != b2)
            {
                length++;
                bytes = b2;
            }

            bw.Write((byte)(0x80 | bytes));
            byte[]? lengthBytes = BitConverter.GetBytes(length);

            for (int i = 0; i < bytes; i++)
            {
                bw.Write(lengthBytes[bytes - i - 1]);
            }
        }
        else
        {
            bw.Write((byte)num);
        }
    }

    private static int GetByteSize(int length) =>
        length < 0xFF
            ? 1
            : length < 0xFFFF
                ? 2
                : length < 0xFFFFFF
                    ? 3
                    : 4;

    private static byte[] GetHashContent(EPPlusSignatureContext ctx, byte[] hash)
    {
        MemoryStream? ms = RecyclableMemory.GetStream();
        BinaryWriter? bw = new BinaryWriter(ms);

        if (ctx.SignatureType == ExcelVbaSignatureType.Legacy)
        {
            bw.Write((byte)0x04); //Octet String Identifier
            bw.Write((byte)hash.Length); //Hash length
            bw.Write(hash); //Content hash
        }
        else
        {
            // SigDataV1Serialized
            bw.Write((byte)0x04); //Octet String Tag Identifier
            const int headerSizeOffset = 4 * 6; // size of header containg size and offset information
            int discriptorLength = headerSizeOffset + ctx.AlgorithmIdentifierOId.Length + 1 + hash.Length; // length of structure
            bw.Write((byte)discriptorLength);
            bw.Write(ctx.AlgorithmIdentifierOId.Length + 1);
            bw.Write(0); // compiled hash size
            bw.Write(hash.Length); // source hash size
            bw.Write(headerSizeOffset); // algorithm id offset
            bw.Write(headerSizeOffset + ctx.AlgorithmIdentifierOId.Length + 1); // compiled hash offset (always empty)
            bw.Write(headerSizeOffset + ctx.AlgorithmIdentifierOId.Length + 1); // source hash offset
            bw.Write(Encoding.ASCII.GetBytes(ctx.AlgorithmIdentifierOId));
            bw.Write((byte)0); // string terminator
            bw.Write(hash);
        }

        bw.Flush();

        return ms.ToArray();
    }

    private static byte GetContentInfoTotalSize() => (byte)0x65;

    private static void WriteOid(BinaryWriter bw, byte[] bytes)
    {
        bw.Write((byte)0x06); //Oid Tag Indentifier 
        bw.Write((byte)bytes.Length); //Lenght OId
        bw.Write(bytes);
    }
}