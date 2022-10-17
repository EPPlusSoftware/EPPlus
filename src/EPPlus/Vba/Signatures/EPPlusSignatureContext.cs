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
        public VbaSignatureHashAlgorithm HashAlgorithm
        {
            get
            {
                switch (AlgorithmIdentifierOId)
                {
                    case HashAlgorithmOids.MD5:
                        return VbaSignatureHashAlgorithm.MD5;
                    case HashAlgorithmOids.SHA256:
                        return VbaSignatureHashAlgorithm.SHA256;
                    case HashAlgorithmOids.SHA384:
                        return VbaSignatureHashAlgorithm.SHA384;
                    case HashAlgorithmOids.SHA512:
                        return VbaSignatureHashAlgorithm.SHA512;
                    default:
                        return VbaSignatureHashAlgorithm.SHA1;
                }
            }
            set
            {
                switch (value)
                {
                    case VbaSignatureHashAlgorithm.MD5:
                        AlgorithmIdentifierOId = HashAlgorithmOids.MD5;
                        break;
                    case VbaSignatureHashAlgorithm.SHA256:
                        AlgorithmIdentifierOId = HashAlgorithmOids.SHA256;
                        break;
                    case VbaSignatureHashAlgorithm.SHA384:
                        AlgorithmIdentifierOId = HashAlgorithmOids.SHA384;
                        break;
                    case VbaSignatureHashAlgorithm.SHA512:
                        AlgorithmIdentifierOId = HashAlgorithmOids.SHA512;
                        break;
                    default:
                        AlgorithmIdentifierOId = HashAlgorithmOids.SHA1;
                        break;
                }
            }
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
            if (string.IsNullOrEmpty(AlgorithmIdentifierOId))
            {
                switch (_signatureType)
                {
                    case ExcelVbaSignatureType.Legacy:
                        AlgorithmIdentifierOId = HashAlgorithmOids.MD5;
                        break;
                    default:
                        AlgorithmIdentifierOId = HashAlgorithmOids.SHA1;
                        break;
                }
            }
            switch (AlgorithmIdentifierOId)
            {
                case HashAlgorithmOids.MD5:
                    return new byte[] { 0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x02, 0x05 };
                case HashAlgorithmOids.SHA1:
                    return new byte[] { 0x2B, 0x0E, 0x03, 0x02, 0x1A };
                case HashAlgorithmOids.SHA256:
                    return new byte[] { 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01 };
                case HashAlgorithmOids.SHA384:
                    return new byte[] { 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x02 };
                case HashAlgorithmOids.SHA512:
                    return new byte[] { 0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x03 };
                default:
                    return null;
            }
        }

        public byte[] GetIndirectDataContentOidBytes()
        {
            switch (_signatureType)
            {
                case ExcelVbaSignatureType.Legacy:
                    return new byte[] { 0x2B, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x1D };
                default:
                    return new byte[] { 0x2B, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x1F };
            }
        }
    }
}
