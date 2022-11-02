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

namespace OfficeOpenXml.VBA.ContentHash
{
    internal static class VbaSignHashAlgorithmUtil
    {
        internal static byte[] GetContentHash(ExcelVbaProject proj, EPPlusSignatureContext ctx)
        {
            if (ctx.SignatureType == ExcelVbaSignatureType.Legacy)
            {
                using (var ms = RecyclableMemory.GetStream())
                {
                    ContentHashInputProvider.GetContentNormalizedDataHashInput(proj, ms);
                    var buffer = ms.ToArray();
                    var hash = ComputeHash(buffer, ctx);
                    var existingHash = ctx.SourceHash;
                    return hash;
                }
            }
            else if (ctx.SignatureType == ExcelVbaSignatureType.Agile)
            {
                using (var ms = RecyclableMemory.GetStream())
                {
                    ContentHashInputProvider.GetContentNormalizedDataHashInput(proj, ms);
                    ContentHashInputProvider.GetFormsNormalizedDataHashInput(proj, ms);
                    var buffer = ms.ToArray();
                    var hash = ComputeHash(buffer, ctx);
                    var existingHash = ctx.SourceHash;
                    return hash;
                }
            }
            else if(ctx.SignatureType == ExcelVbaSignatureType.V3)
            {
                using (var ms = RecyclableMemory.GetStream())
                {
                    ContentHashInputProvider.GetV3ContentNormalizedDataHashInput(proj, ms);
                    var buffer = ms.ToArray();
                    var hash = ComputeHash(buffer, ctx);
                    var existingHash = ctx.SourceHash;
                    
                    return hash;
                }
            }
            return default(byte[]);
            
        }
        internal static byte[] ComputeHash(byte[] buffer, EPPlusSignatureContext ctx)
        {
            var algorithm = ctx.GetHashAlgorithm();
            if (algorithm == null) return null;
            return algorithm.ComputeHash(buffer);
        }
    }
}
