using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Encryption;
using System;
using System.Security.Cryptography;

namespace OfficeOpenXml.Utils.EncodingUtils
{
    internal class EncodeUtil
    {
        public static string HashAndEncodeBytes(byte[] temp)
        {
            using (var sha1Hash = SHA1.Create())
            {
                var hash = sha1Hash.ComputeHash(temp);
                return Convert.ToBase64String(hash);
            }
        }

        public static HashAlgorithm GetHashProvider(DigitalSignatureHashAlgorithm signatureHashAlgorithm)
        {
            switch (signatureHashAlgorithm)
            {
                case DigitalSignatureHashAlgorithm.SHA1:
                    return SHA1.Create();
                case DigitalSignatureHashAlgorithm.SHA256:
                    return SHA256.Create();
                case DigitalSignatureHashAlgorithm.SHA384:
                    return SHA384.Create();
                case DigitalSignatureHashAlgorithm.SHA512:
                    return SHA512.Create();
                default:
                    throw new NotSupportedException(string.Format("Hash provider is unsupported. {0}", signatureHashAlgorithm));
            }
        }

        public static string HashAndEncodeBytes(byte[] temp, DigitalSignatureHashAlgorithm hashAlgorithm)
        {
            using (var hashProvider = GetHashProvider(hashAlgorithm))
            {
                var hash = hashProvider.ComputeHash(temp);
                return Convert.ToBase64String(hash);
            }
        }
    }
}
