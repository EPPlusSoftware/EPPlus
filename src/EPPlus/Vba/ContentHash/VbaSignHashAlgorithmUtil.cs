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
            #region old code
            ////MS-OVBA 2.4.2
            //var enc = System.Text.Encoding.GetEncoding(proj.CodePage);
            //using (var ms = RecyclableMemory.GetStream())
            //{
            //    BinaryWriter bw = new BinaryWriter(ms);
            //    bw.Write(enc.GetBytes(proj.Name));
            //    bw.Write(enc.GetBytes(proj.Constants));
            //    foreach (var reference in proj.References)
            //    {
            //        if (reference.ReferenceRecordID == 0x0D)
            //        {
            //            bw.Write((byte)0x7B);
            //        }
            //        else if (reference.ReferenceRecordID == 0x0E)
            //        {
            //            foreach (byte b in BitConverter.GetBytes((uint)reference.Libid.Length))  //Length will never be an UInt with 4 bytes that aren't 0 (> 0x00FFFFFF), so no need for the rest of the properties.
            //            {
            //                if (b != 0)
            //                {
            //                    bw.Write(b);
            //                }
            //                else
            //                {
            //                    break;
            //                }
            //            }
            //        }
            //    }
            //    foreach (var module in proj.Modules)
            //    {
            //        var lines = module.Code.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            //        foreach (var line in lines)
            //        {
            //            if (!line.StartsWith("attribute", StringComparison.OrdinalIgnoreCase))
            //            {
            //                bw.Write(enc.GetBytes(line));
            //            }
            //        }
            //    }
            #endregion

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
                    Debug_WriteInfo(proj, buffer);
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
                    File.WriteAllBytes(@"c:\epplusTest\testoutput\v3Signing\v3sign.bin", buffer);
                    var hash = ComputeHash(buffer, ctx);
                    var existingHash = ctx.SourceHash;
                    
                    return hash;
                }
            }
            return default(byte[]);
            
        }

        private static void Debug_WriteInfo(ExcelVbaProject proj, byte[] buffer)
        {
            var sw = new StringWriter();
            foreach (var d in proj.Document.Directories)
            {
                sw.WriteLine(d.FullName);
            }
            File.WriteAllText("c:\\temp\\documents.txt", sw.ToString());
            File.WriteAllBytes("c:\\temp\\agile.bin", buffer);
        }

        internal static byte[] ComputeHash(byte[] buffer, EPPlusSignatureContext ctx)
        {
            var algorithm = ctx.GetHashAlgorithm();
            if (algorithm == null) return null;
            return algorithm.ComputeHash(buffer);
        }
    }
}
