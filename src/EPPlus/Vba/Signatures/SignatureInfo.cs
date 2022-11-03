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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace OfficeOpenXml.VBA.Signatures
{
    internal class SignatureInfo
    {
        public uint cbSignature;
        public uint signatureOffset;     //44 ??
        public uint cbSigningCertStore;
        public uint certStoreOffset;
        public uint cbProjectName;
        public uint projectNameOffset;
        public uint fTimestamp;
        public uint cbTimestampUrl;
        public uint timestampUrlOffset;
        public byte[] signature;
        public uint version;
        public uint fileType;

        public uint id;
        internal uint endel1;
        internal uint endel2;
        internal ushort rgchProjectNameBuffer;
        internal ushort rgchTimestampBuffer;

        public X509Certificate2 Certificate { get; internal set; }
        public SignedCms Verifier { get; internal set; }
    }
}
