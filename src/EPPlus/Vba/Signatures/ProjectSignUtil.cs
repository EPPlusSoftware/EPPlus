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
using OfficeOpenXml.VBA.ContentHash;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.Pkcs;
using System.Text;

namespace OfficeOpenXml.VBA.Signatures
{
    internal static class ProjectSignUtil
    {
        internal static ContentInfo SignProject(ExcelVbaProject proj, EPPlusVbaSignature signature, EPPlusSignatureContext ctx)
        {
            var certificate = signature.Certificate;
            if (!certificate.HasPrivateKey)
            {
                //throw (new InvalidOperationException("The certificate doesn't have a private key"));
                signature.Certificate = null;
                return null;
            }

            var hash = VbaSignHashAlgorithmUtil.GetContentHash(proj, ctx);

            ContentInfo contentInfo;
            using (var ms = RecyclableMemory.GetStream())
            {
                contentInfo = CreateContentInfo(hash, ms);
            }
            contentInfo.ContentType.Value = "1.3.6.1.4.1.311.2.1.4";
            return contentInfo;
        }

        private static ContentInfo CreateContentInfo(byte[] hash, MemoryStream ms)
        {
            ContentInfo contentInfo;
            // [MS-OSHARED] 2.3.2.4.3.1 SpcIndirectDataContent
            BinaryWriter bw = new BinaryWriter(ms);
            bw.Write((byte)0x30); //Constructed Type 
            bw.Write((byte)0x32); //Total length
            bw.Write((byte)0x30); //Constructed Type 
            bw.Write((byte)0x0E); //Length SpcIndirectDataContent
            bw.Write((byte)0x06); //Oid Tag Indentifier 
            bw.Write((byte)0x0A); //Lenght OId
            bw.Write(new byte[] { 0x2B, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x1D }); //Encoded Oid 1.3.6.1.4.1.311.2.1.29
            bw.Write((byte)0x04);   //Octet String Tag Identifier
            bw.Write((byte)0x00);   //Zero length

            bw.Write((byte)0x30); //Constructed Type (DigestInfo)
            bw.Write((byte)0x20); //Length DigestInfo
            bw.Write((byte)0x30); //Constructed Type (Algorithm)
            bw.Write((byte)0x0C); //length AlgorithmIdentifier
            bw.Write((byte)0x06); //Oid Tag Indentifier 
            bw.Write((byte)0x08); //Lenght OId

            // SHA1: 42, 134, 72, 134, 247, 13, 2, 5
            bw.Write(new byte[] { 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x02, 0x05 }); //Encoded Oid for 1.2.840.113549.2.5 (AlgorithmIdentifier MD5)
            bw.Write((byte)0x05);   //Null type identifier
            bw.Write((byte)0x00);   //Null length
            bw.Write((byte)0x04);   //Octet String Identifier
            bw.Write((byte)hash.Length);   //Hash length
            bw.Write(hash);                //Content hash

            contentInfo = new ContentInfo(ms.ToArray());
            return contentInfo;
        }
    }
}
