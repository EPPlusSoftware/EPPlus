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
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace OfficeOpenXml.VBA.Signatures
{
    internal static class SignatureReader
    {
        private const string IndirectDataContentOid   = "1.3.6.1.4.1.311.2.1.29";
        private const string IndirectDataContentOidV2 = "1.3.6.1.4.1.311.2.1.31";
        internal static SignatureInfo ReadSignature(ZipPackagePart part, ExcelVbaSignatureType signatureType, EPPlusSignatureContext ctx)
        {
            // [MS-OSHARED] 2.3.2.1 DigSigInfoSerialized
            var si = new SignatureInfo();
            var stream = part.GetStream();
            BinaryReader br = new BinaryReader(stream);
            si.cbSignature = br.ReadUInt32();
            si.signatureOffset = br.ReadUInt32();     //44 ??
            si.cbSigningCertStore = br.ReadUInt32();
            si.certStoreOffset = br.ReadUInt32();
            si.cbProjectName = br.ReadUInt32();
            si.projectNameOffset = br.ReadUInt32();
            si.fTimestamp = br.ReadUInt32();
            si.cbTimestampUrl = br.ReadUInt32();
            si.timestampUrlOffset = br.ReadUInt32();
            si.signature = br.ReadBytes((int)si.cbSignature);

            //Read serialized Properties MS-OSHARED 2.3.2.5.5 VBASigSerializedCertStore
            si.version = br.ReadUInt32();
            si.fileType = br.ReadUInt32();

            uint id = br.ReadUInt32();
            while (id != 0)
            {
                uint encodingType = br.ReadUInt32();
                uint length = br.ReadUInt32();
                if (length > 0)
                {
                    byte[] value = br.ReadBytes((int)length);
                    switch (id)
                    {
                        //Add property values here...
                        case 0x20:
                            si.Certificate = new X509Certificate2(value);
                            break;
                        default:
                            break;
                    }
                }
                id = br.ReadUInt32();
            }
            si.endel1 = br.ReadUInt32();  //0
            si.endel2 = br.ReadUInt32();  //0
            si.rgchProjectNameBuffer = br.ReadUInt16();
            si.rgchTimestampBuffer = br.ReadUInt16();

            si.Verifier = new SignedCms();
            si.Verifier.Decode(si.signature);
            ReadSignedData(si.Verifier.ContentInfo.Content, ctx);
            return si;
        }

        internal static void ReadSignedData(byte[] data, EPPlusSignatureContext ctx)
        {
            var ms = RecyclableMemory.GetStream(data);
            var br = new BinaryReader(ms);            
            var totallength = ReadSequence(br);
            var lengthSpcIndirectDataContent = ReadSequence(br);
            var indirectDataContentOid = ReadOId(br);
            var digestValue = ReadOctStringBytes(br);

            var lengthDigestInfo = ReadSequence(br);
            var lengthAlgorithmIdentifier = ReadSequence(br);
            ctx.AlgorithmIdentifierOId = ReadOId(br);

            //Parameter is null
            var nullTypeIdentifyer = br.ReadByte();   //Null type identifier
            var nullLength = br.ReadByte();   //Null length

            if (indirectDataContentOid == IndirectDataContentOidV2) //V2
            {
                //Read
                var SigFormatDescriptorV1_size = BitConverter.ToInt32(digestValue, 0);    //12
                var SigFormatDescriptorV1_version = BitConverter.ToInt32(digestValue, 4); //1
                var SigFormatDescriptorV1_format = BitConverter.ToInt32(digestValue, 8);  //1

                //var sigDataV1Serialized = ReadOctStringBytes(br); //SigDataV1Serialized
                var id = br.ReadByte();  //4
                var octstringSize = br.ReadByte();
                var sigDataV1Serialized_algorithmIdSize = br.ReadInt32();
                var sigDataV1Serialized_compiledHashSize = br.ReadInt32();
                var sigDataV1Serialized_sourceHashSize = br.ReadInt32();
                var sigDataV1Serialized_algorithmIdOffset = br.ReadInt32();
                var sigDataV1Serialized_compiledHashOffset = br.ReadInt32();
                var sigDataV1Serialized_sourceHashOffset = br.ReadInt32();

                var sigDataV1Serialized_algorithmId = br.ReadBytes(sigDataV1Serialized_algorithmIdSize);    //As a string here apparently. Should match the AlgorithmIdentifierOId above.
                var algId = Encoding.ASCII.GetString(sigDataV1Serialized_algorithmId, 0, sigDataV1Serialized_algorithmIdSize - 1); //Skip ending \0
                var sigDataV1Serialized_compiledHash = br.ReadBytes(sigDataV1Serialized_compiledHashSize);
                var sigDataV1Serialized_sourceHash = br.ReadBytes(sigDataV1Serialized_sourceHashSize); //ReadOctStringBytes(br);
                ctx.AlgorithmIdentifierOId = algId;
                ctx.CompiledHash = sigDataV1Serialized_compiledHash;
                ctx.SourceHash = sigDataV1Serialized_sourceHash;
            }
            else  //V1
            {
                var hash = ReadOctStringBytes(br);
                ctx.SourceHash = hash;
            }
        }

        private static int ReadSequence(BinaryReader br)
        {
            var id = br.ReadByte();
            if (id == 0x30)
            {
                var b = br.ReadByte();
                if (b > 0x80)
                {
                    var bl = b & 0x80;
                    var lengthBytes = br.ReadBytes(bl);
                    return BitConverter.ToInt32(lengthBytes.Reverse().ToArray(), 0);
                }                
                return b;
            }
            return id;
        }

        private static byte[] ReadOctStringBytes(BinaryReader bw)
        {
            var s = "";
            var id = bw.ReadByte();   //Octet String Tag Identifier
            if (id == 4)
            {
                var octetStringLength = bw.ReadByte();   //Zero length

                if (octetStringLength > 0)
                {
                    return bw.ReadBytes(octetStringLength);
                }
            }
            return default(byte[]);
        }

        //Create Oid from a bytearray
        internal static string ReadHash(byte[] content, int offset = 6)
        {
            StringBuilder builder = new StringBuilder();
            //int offset = 0x6;
            if (0 < content.Length)
            {
                byte num = content[offset];
                byte num2 = (byte)(num / 40);
                builder.Append(num2.ToString(null, null));
                builder.Append(".");
                num2 = (byte)(num % 40);
                builder.Append(num2.ToString(null, null));
                ulong num3 = 0L;
                for (int i = offset + 1; i < content.Length; i++)
                {
                    num2 = content[i];
                    num3 = (num3 << 7) + (byte)(num2 & 0x7f);
                    if ((num2 & 0x80) == 0)
                    {
                        builder.Append(".");
                        builder.Append(num3.ToString(null, null));
                        num3 = 0L;
                    }
                    //1.2.840.113549.2.5
                }
            }


            string oId = builder.ToString();

            return oId;
        }

        internal static string ReadOId(BinaryReader bw)
        {
            var oIdIdentifyer = bw.ReadByte();
            if (oIdIdentifyer == 6)
            {
                var length = bw.ReadByte();
                var oidData = bw.ReadBytes(length);
                return ReadHash(oidData, 0);
            }
            return null;
        }
    }
}
