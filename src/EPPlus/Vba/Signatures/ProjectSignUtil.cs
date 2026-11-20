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
using System.Security.Cryptography;
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
                signature.Certificate = null;
                return null;
            }

            var hash = VbaSignHashAlgorithmUtil.GetContentHash(proj, ctx);

            return  CreateContentInfo(hash, ctx);
        }
        private static ContentInfo CreateContentInfo(byte[] hash, EPPlusSignatureContext ctx)
        {
            // [MS-OSHARED] 2.3.2.4.3.1 SpcIndirectDataContent
            //   SEQUENCE {
            //     SEQUENCE { OID spcIndirectDataContent / SpcAttributeType, OCTET STRING (empty | SigFormatDescriptorV1) }
            //     DigestInfo SEQUENCE { AlgorithmIdentifier SEQUENCE { OID, NULL }, OCTET STRING (hash | SigDataV1Serialized) }
            //   }
            var attributeValue = ctx.SignatureType == ExcelVbaSignatureType.Legacy
                ? new byte[0]
                : GetSigFormatDescriptorV1();

            var data = Sequence(
                Oid(ctx.GetIndirectDataContentOidBytes()),
                OctetString(attributeValue));

            var digestInfo = Sequence(
                Sequence(Oid(ctx.GetHashAlgorithmBytes()), DerNull),
                OctetString(GetDigestValue(ctx, hash)));

            var spcIndirectDataContent = Sequence(data, digestInfo);
            return new ContentInfo(new Oid("1.3.6.1.4.1.311.2.1.4"), spcIndirectDataContent);
        }

        private static byte[] GetSigFormatDescriptorV1()
        {
            using (var ms = EPPlusMemoryManager.GetStream())
            {
                var bw = new BinaryWriter(ms);
                bw.Write(12); // size
                bw.Write(1);  // version
                bw.Write(1);  // format
                bw.Flush();
                return ms.ToArray();
            }
        }

        // Returns the raw OCTET STRING value (no tag/length).
        private static byte[] GetDigestValue(EPPlusSignatureContext ctx, byte[] hash)
        {
            if (ctx.SignatureType == ExcelVbaSignatureType.Legacy)
                return hash;

            // SigDataV1Serialized
            using (var ms = EPPlusMemoryManager.GetStream())
            {
                var bw = new BinaryWriter(ms);
                const int headerSize = 4 * 6;
                var algId = Encoding.ASCII.GetBytes(ctx.AlgorithmIdentifierOId + "\0");

                bw.Write(algId.Length);              // algorithmIdSize
                bw.Write(0);                         // compiledHashSize
                bw.Write(hash.Length);               // sourceHashSize
                bw.Write(headerSize);                // algorithmIdOffset
                bw.Write(headerSize + algId.Length); // compiledHashOffset
                bw.Write(headerSize + algId.Length); // sourceHashOffset
                bw.Write(algId);
                bw.Write(hash);
                bw.Flush();
                return ms.ToArray();
            }
        }

        #region Minimal DER writer
        private static readonly byte[] DerNull = { 0x05, 0x00 };

        private static byte[] Sequence(params byte[][] children) => Tlv(0x30, children);
        private static byte[] Oid(byte[] encodedOid) => Tlv(0x06, encodedOid);
        private static byte[] OctetString(byte[] value) => Tlv(0x04, value);

        private static byte[] Tlv(byte tag, params byte[][] children)
        {
            int length = 0;
            foreach (var c in children) length += c.Length;

            using (var ms = new MemoryStream())
            {
                ms.WriteByte(tag);
                WriteDerLength(ms, length);
                foreach (var c in children) ms.Write(c, 0, c.Length);
                return ms.ToArray();
            }
        }

        private static void WriteDerLength(Stream s, int length)
        {
            if (length < 0x80)
            {
                s.WriteByte((byte)length);
                return;
            }
            int n = length > 0xFFFFFF ? 4 : length > 0xFFFF ? 3 : length > 0xFF ? 2 : 1;
            s.WriteByte((byte)(0x80 | n));
            for (int i = n - 1; i >= 0; i--)
                s.WriteByte((byte)(length >> (8 * i)));
        }
        #endregion
    }
}
