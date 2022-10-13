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
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace OfficeOpenXml.VBA.Signatures
{
    internal static class CertUtil
    {
        internal static X509Certificate2 GetCertFromStore(StoreLocation loc, string thumbPrint)
        {
            try
            {
                X509Store store = new X509Store(StoreName.My, loc);
                store.Open(OpenFlags.ReadOnly);
                try
                {
                    var storeCert = store.Certificates.Find(
                                    X509FindType.FindByThumbprint,
                                    thumbPrint,
                                    false
                                    ).OfType<X509Certificate2>().FirstOrDefault();
                    return storeCert;
                }
                finally
                {
#if Core
                    store.Dispose();
#endif
                    store.Close();
                }
            }
            catch
            {
                return null;
            }
        }

        internal static byte[] GetSerializedCertStore(byte[] certRawData)
        {
            //MS-OSHARED 2.3.2.5.5 VBASigSerializedCertStore
            using (var ms = RecyclableMemory.GetStream())
            {
                var bw = new BinaryWriter(ms);

                bw.Write((uint)0); //Version
                bw.Write((uint)0x54524543); //fileType

                //SerializedCertificateEntry
                var certData = certRawData;
                bw.Write((uint)0x20);
                bw.Write((uint)1);
                bw.Write((uint)certData.Length);
                bw.Write(certData);

                //EndElementMarkerEntry
                bw.Write((uint)0);
                bw.Write((ulong)0);

                bw.Flush();
                return ms.ToArray();
            }
        }

        internal static byte[] CreateBinarySignature(MemoryStream ms, BinaryWriter bw, byte[] certStore, byte[] cert)
        {
            // [MS-OSHARED] 2.3.2.1 DigSigInfoSerialized
            bw.Write((uint)cert.Length);
            bw.Write((uint)44);                  //?? 36 ref inside cert ??
            bw.Write((uint)certStore.Length);    //cbSigningCertStore
            bw.Write((uint)(cert.Length + 44));  //certStoreOffset
            bw.Write((uint)0);                   //cbProjectName
            bw.Write((uint)(cert.Length + certStore.Length + 44));    //projectNameOffset
            bw.Write((uint)0);    //fTimestamp
            bw.Write((uint)0);    //cbTimestampUrl
            bw.Write((uint)(cert.Length + certStore.Length + 44 + 2));    //timestampUrlOffset
            bw.Write(cert);
            bw.Write(certStore);
            bw.Write((ushort)0);//rgchProjectNameBuffer
            bw.Write((ushort)0);//rgchTimestampBuffer
            bw.Write((ushort)0);
            bw.Flush();
            var b = ms.ToArray();
            return b;
        }

        internal static X509Certificate2 GetCertificate(string thumbprint)
        {
            var storeCert = GetCertFromStore(StoreLocation.CurrentUser, thumbprint);
            if (storeCert == null)
            {
                storeCert = GetCertFromStore(StoreLocation.LocalMachine, thumbprint);
            }
            if (storeCert != null && storeCert.HasPrivateKey == true)
            {
                return storeCert;
            }
            return null;
        }

        internal static SignedCms SignProject(ExcelVbaProject proj, EPPlusVbaSignature signature, EPPlusSignatureContext ctx)
        {
            var contentInfo = ProjectSignUtil.SignProject(proj, signature, ctx);
            File.WriteAllBytes(@"c:\Temp\contentInfoWrite.bin", contentInfo.Content);
            var verifier = new SignedCms(contentInfo);
            var signer = new CmsSigner(signature.Certificate);
            verifier.ComputeSignature(signer, false);
            return verifier;
        }
    }
}
