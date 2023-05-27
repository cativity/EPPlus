/*************************************************************************************************
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
using OfficeOpenXml.Vba.ContentHash;
using OfficeOpenXml.VBA.Signatures;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace OfficeOpenXml.VBA.ContentHash;

internal static class VbaSignHashAlgorithmUtil
{
    internal static byte[] GetContentHash(ExcelVbaProject proj, EPPlusSignatureContext ctx)
    {
        if (ctx.SignatureType == ExcelVbaSignatureType.Legacy)
        {
            using MemoryStream? ms = RecyclableMemory.GetStream();
            ContentHashInputProvider.GetContentNormalizedDataHashInput(proj, ms);
            byte[]? buffer = ms.ToArray();
            byte[]? hash = ComputeHash(buffer, ctx);

            return hash;
        }
        else if (ctx.SignatureType == ExcelVbaSignatureType.Agile)
        {
            using MemoryStream? ms = RecyclableMemory.GetStream();
            ContentHashInputProvider.GetContentNormalizedDataHashInput(proj, ms);
            ContentHashInputProvider.GetFormsNormalizedDataHashInput(proj, ms);
            byte[]? buffer = ms.ToArray();
            byte[]? hash = ComputeHash(buffer, ctx);

            return hash;
        }
        else if (ctx.SignatureType == ExcelVbaSignatureType.V3)
        {
            using MemoryStream? ms = RecyclableMemory.GetStream();
            ContentHashInputProvider.GetV3ContentNormalizedDataHashInput(proj, ms);
            byte[]? buffer = ms.ToArray();
            byte[]? hash = ComputeHash(buffer, ctx);

            return hash;
        }

        return default(byte[]);
    }

    internal static byte[] ComputeHash(byte[] buffer, EPPlusSignatureContext ctx)
    {
        HashAlgorithm? algorithm = ctx.GetHashAlgorithm();

        if (algorithm == null)
        {
            return null;
        }

        return algorithm.ComputeHash(buffer);
    }
}