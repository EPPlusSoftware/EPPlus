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
using System.Security.Cryptography;
using System.Text;

namespace OfficeOpenXml.VBA.Signatures
{
    internal class EPPlusSignatureContext
    {
        private static class HashAlgorithmOids
        {
            public const string MD5 = "1.2.840.113549.2.5";
            public const string SHA1 = "1.3.14.3.2.26";
            public const string SHA256 = "2.16.840.1.101.3.4.2.1";
            public const string SHA384 = "2.16.840.1.101.3.4.2.2";
            public const string SHA512 = "2.16.840.1.101.3.4.2.3";
        }
        public EPPlusSignatureContext(ExcelVbaSignatureType signatureType)
        {
            _signatureType = signatureType;
        }

        private readonly ExcelVbaSignatureType _signatureType;

        public ExcelVbaSignatureType SignatureType => _signatureType;

        public string AlgorithmIdentifierOId
        {
            get;
            set;
        }

        public byte[] CompiledHash
        {
            get;
            set;
        }

        public byte[] SourceHash
        {
            get;
            set;
        }

        public HashAlgorithm GetHashAlgorithm()
        {
            if (string.IsNullOrEmpty(AlgorithmIdentifierOId)) return GetHashAlgorithmDefault();
            switch(AlgorithmIdentifierOId)
            {
                case HashAlgorithmOids.MD5:
                    return MD5.Create();
                case "1.3.14.3.2.26":
                    return SHA1.Create();
                case "2.16.840.1.101.3.4.2.1":
                    return SHA256.Create();
                case "2.16.840.1.101.3.4.2.2":
                    return SHA384.Create();
                case "2.16.840.1.101.3.4.2.3":
                    return SHA512.Create();
                default:
                    return null;
            }
        }

        private HashAlgorithm GetHashAlgorithmDefault()
        {
            switch(_signatureType)
            {
                case ExcelVbaSignatureType.Legacy:
                    return MD5.Create();
                default:
                    return SHA1.Create();
            }
        }

        public byte[] GetHashAlgorithmBytes()
        {
            switch (AlgorithmIdentifierOId)
            {
                case HashAlgorithmOids.MD5:
                    return new byte[] { 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x02, 0x05 };
                case HashAlgorithmOids.SHA1:
                    return new byte[] { 0x2B, 0x0E, 0x03, 0x02, 0x1A };
                default:
                    return null;
            }
        }
    }
}
