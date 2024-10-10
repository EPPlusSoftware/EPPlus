using System;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Security.Cryptography;
using System.Xml;
using System.Globalization;
using System.Linq;

namespace OfficeOpenXml.DigitalSignatures.XAdES
{
    internal class SignedSignatureProperites
    {
        DateTime SigningTime;
        string Prefix = "xd";
        string Algorithm = "http://www.w3.org/2000/09/xmldsig#sha1";
        string Name;
        internal string Serial;
        X509Certificate2 Cert;
        string Hash = null;
        string TimeStr = null;
        AdditionalSignatureInfo _info;

        internal static string BytesToNumericString(byte[] bytes)
        {
            //minimum array length should be 1
            if (bytes.Length == 0)
            {
                return "0";
            }

            //the logarithm of 2 in base 10 (Or in this case think of exponent of 10 to become 2)
            var logOf2 = Math.Log(2, 10);
            var empty16Bits = 0x80000;
            var full16Bits = 0xFFFF;
            var constant = (int)(logOf2 * empty16Bits);

            var numBytes = bytes.Length;
            var bytesMultConstant = numBytes * constant;
            var maxDigits = (bytesMultConstant + full16Bits) >> 16;

            var digitsArray = new byte[maxDigits];
            int len = 1;

            for (int j = 0; j != bytes.Length; ++j)
            {
                int i;
                int carriedOverDigits = bytes[j];
                for (i = 0; i < len || carriedOverDigits != 0; i++)
                {
                    //If there's already data in the digit we are trying to write to
                    //Offset value forward by that much. Since one byte is a value between 0 and 255
                    //Essentially we are performing the shift of multiplying by '10' but in byte values 256
                    var byteOffset = digitsArray[i] * 256;

                    int currentDigitValue = byteOffset + carriedOverDigits;
                    carriedOverDigits = Math.DivRem(currentDigitValue, 10, out currentDigitValue);
                    digitsArray[i] = (byte)currentDigitValue;
                }
                //We filled another digit and made length longer
                if (i > len)
                {
                    len = i;
                }
            }

            var numString = new StringBuilder(len);
            while (len > 0)
            {
                len = len - 1;
                numString.Append((char)('0' + digitsArray[len]));
            }
            return numString.ToString();
        }

        public string HashAndEncodeBytes(byte[] temp)
        {
            using (var sha1Hash = SHA1.Create())
            {
                var hash = sha1Hash.ComputeHash(temp);
                return Convert.ToBase64String(hash);
            }
        }

        internal SignedSignatureProperites(string prefix, XmlElement SignedSignaturePropertiesNode, AdditionalSignatureInfo info)
        {
            Prefix = prefix;
            TimeStr = SignedSignaturePropertiesNode.GetElementsByTagName($"{prefix}:SigningTime")[0].InnerText;
            Algorithm = SignedSignaturePropertiesNode.SelectSingleNode("//*[@Algorithm]").Attributes.GetNamedItem("Algorithm").InnerText;
            Hash = SignedSignaturePropertiesNode.GetElementsByTagName("DigestValue")[0].InnerText;
            Name = SignedSignaturePropertiesNode.GetElementsByTagName("X509IssuerName")[0].InnerText;
            Serial = SignedSignaturePropertiesNode.GetElementsByTagName("X509SerialNumber")[0].InnerText;
            var prodPlace = (XmlElement) SignedSignaturePropertiesNode.GetElementsByTagName($"{prefix}:SignatureProductionPlace")[0];

            if(prodPlace != null)
            {
                var cityNode = prodPlace.GetElementsByTagName($"{prefix}:City")[0];
                if(cityNode != null)
                {
                    info.City = string.IsNullOrEmpty(cityNode.InnerText) ? null : cityNode.InnerText;
                }

                var stateNode = prodPlace.GetElementsByTagName($"{prefix}:StateOrProvince")[0];
                if (stateNode != null)
                {
                    info.StateOrProvince = string.IsNullOrEmpty(stateNode.InnerText) ? null : stateNode.InnerText;
                }
                var postalCode = prodPlace.GetElementsByTagName($"{prefix}:PostalCode")[0];
                if (postalCode != null)
                {
                    info.ZIPorPostalCode = string.IsNullOrEmpty(postalCode.InnerText) ? null : postalCode.InnerText;
                }

                var countryName = prodPlace.GetElementsByTagName($"{prefix}:CountryName")[0];
                if (countryName != null)
                {
                    info.CountryOrRegion = string.IsNullOrEmpty(countryName.InnerText) ? null : countryName.InnerText;
                }
            }

            var roleNode = SignedSignaturePropertiesNode.GetElementsByTagName($"{prefix}:ClaimedRole")[0];
            if(roleNode != null)
            {
                info.SignerRoleTitle = string.IsNullOrEmpty(roleNode.InnerText) ? null : roleNode.InnerText;
            }

            _info = info;
        }

        internal SignedSignatureProperites(string prefix, X509Certificate2 cert, AdditionalSignatureInfo info)
        {
            Prefix = prefix;
            SigningTime = DateTime.Now;
            TimeStr = SigningTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
            Cert = cert;
            Name = cert.Issuer;
            Hash = HashAndEncodeBytes(Cert.RawData);
            _info = info;

            var bytes = Cert.GetSerialNumber();
            bytes = bytes.Reverse().ToArray();
            Serial = BytesToNumericString(bytes);
        }

        internal string GetXML()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append($"<{Prefix}:SignedSignatureProperties>");

            sb.Append($"<{Prefix}:SigningTime>{TimeStr}</{Prefix}:SigningTime>");

            sb.Append($"<{Prefix}:SigningCertificate>");
            sb.Append($"<{Prefix}:Cert>");

            sb.Append($"<{Prefix}:CertDigest>");
            sb.Append($"<DigestMethod Algorithm=\"{Algorithm}\"></DigestMethod>");
            sb.Append($"<DigestValue>{Hash}</DigestValue>");
            sb.Append($"</{Prefix}:CertDigest>");

            sb.Append($"<{Prefix}:IssuerSerial>");
            sb.Append($"<X509IssuerName>{Name}</X509IssuerName>");
            sb.Append($"<X509SerialNumber>{Serial}</X509SerialNumber>");
            sb.Append($"</{Prefix}:IssuerSerial>");

            sb.Append($"</{Prefix}:Cert>");
            sb.Append($"</{Prefix}:SigningCertificate>");

            sb.Append($"<{Prefix}:SignaturePolicyIdentifier>");
            sb.Append($"<{Prefix}:SignaturePolicyImplied></{Prefix}:SignaturePolicyImplied>");
            sb.Append($"</{Prefix}:SignaturePolicyIdentifier>");

            string signatureProductionPlace = $"<{Prefix}:SignatureProductionPlace>";

            if (_info.City != null)
            {
                signatureProductionPlace += $"<{Prefix}:City>{_info.City}</{Prefix}:City>";
            }
            if (_info.StateOrProvince != null)
            {
                signatureProductionPlace += $"<{Prefix}:StateOrProvince>{_info.StateOrProvince}</{Prefix}:StateOrProvince>";
            }
            if (_info.ZIPorPostalCode != null)
            {
                signatureProductionPlace += $"<{Prefix}:PostalCode>{_info.ZIPorPostalCode}</{Prefix}:PostalCode>";
            }
            if (_info.CountryOrRegion != null)
            {
                signatureProductionPlace += $"<{Prefix}:CountryName>{_info.CountryOrRegion}</{Prefix}:CountryName>";
            }

            if(signatureProductionPlace != $"<{Prefix}:SignatureProductionPlace>")
            {
                signatureProductionPlace += $"</{Prefix}:SignatureProductionPlace>";
                sb.Append(signatureProductionPlace);
            }

            if(_info.SignerRoleTitle != null) 
            {
                sb.Append
                    ($"<{Prefix}:SignerRole>" +
                         $"<{Prefix}:ClaimedRoles>" +
                            $"<{Prefix}:ClaimedRole>" +
                                $"{_info.SignerRoleTitle}" +
                            $"</{Prefix}:ClaimedRole>" +
                        $"</{Prefix}:ClaimedRoles>" + 
                    $"</{Prefix}:SignerRole>");
            }

            sb.Append($"</{Prefix}:SignedSignatureProperties>");

            return sb.ToString();
        }
    }
}
