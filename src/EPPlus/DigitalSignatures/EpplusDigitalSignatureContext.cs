using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;

namespace OfficeOpenXml.DigitalSignatures
{
    internal static class DigestMethods
    {
        public const string SHA1 = "http://www.w3.org/2000/09/xmldsig#sha1";
        public const string SHA256 = "http://www.w3.org/2001/04/xmlenc#sha256";
        public const string SHA384 = "http://www.w3.org/2001/04/xmlenc#sha384";
        public const string SHA512 = "http://www.w3.org/2001/04/xmlenc#sha512";

        internal static string GetDigestMethod(DigitalSignatureHashAlgorithm algorithm)
        {
            switch (algorithm)
            {
                case DigitalSignatureHashAlgorithm.SHA1:
                    return SHA1;
                case DigitalSignatureHashAlgorithm.SHA256:
                    return SHA256;
                case DigitalSignatureHashAlgorithm.SHA384:
                    return SHA384;
                case DigitalSignatureHashAlgorithm.SHA512:
                    return SHA512;
                default:
                    throw new InvalidOperationException($"The hash algorithm '{algorithm}' is invalid. Please use another algorithm");
            }
        }

        internal static DigitalSignatureHashAlgorithm? GetHashAlgorithmByDigest(string digestMethod)
        {
            switch (digestMethod)
            {
                case SHA1:
                    return DigitalSignatureHashAlgorithm.SHA1;
                case SHA256:
                    return DigitalSignatureHashAlgorithm.SHA256;
                case SHA384:
                    return DigitalSignatureHashAlgorithm.SHA384;
                case SHA512:
                    return DigitalSignatureHashAlgorithm.SHA512;
                default:
                    return null;
                    throw new InvalidOperationException($"The string '{digestMethod}' is undefined. It does not match any valid DigitalSignatureHashAlgorithm. Please use a defined string");
            }
        }

        internal static string GetSignatureMethod(DigitalSignatureHashAlgorithm algorithm)
        {
            switch (algorithm)
            {
                case DigitalSignatureHashAlgorithm.SHA1:
                    return "http://www.w3.org/2000/09/xmldsig#rsa-sha1";
                case DigitalSignatureHashAlgorithm.SHA256:
                    return "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256";
                case DigitalSignatureHashAlgorithm.SHA384:
                    return "http://www.w3.org/2001/04/xmldsig-more#rsa-sha384";
                case DigitalSignatureHashAlgorithm.SHA512:
                    return "http://www.w3.org/2001/04/xmldsig-more#rsa-sha512";
                default:
                    throw new InvalidOperationException($"The hash algorithm '{algorithm}' is invalid. Please use another algorithm");
            }
        }
    }

    internal class EpplusDigitalSignatureContext
    {

    }
}
