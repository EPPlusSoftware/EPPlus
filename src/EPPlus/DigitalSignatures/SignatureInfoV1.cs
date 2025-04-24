using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.DigitalSignatures
{
    /// <summary>
    /// Information about the signature including OSversion and office version
    /// </summary>
    internal class SignatureInfoV1
    {
        internal SignatureInfoV1(bool eastAsianProvider = false) 
        {
            string defaultSignatureProvider = "{00000000-0000-0000-0000-000000000000}";
            _signatureProviderID = defaultSignatureProvider;
            _windowsVersion = Environment.OSVersion.Version.ToString();

            if (eastAsianProvider)
            {
                _signatureProviderID = "{000CD6A4-0000-0000-C000-000000000046}";
            }
        }

        internal SignatureInfoV1(XmlElement signatureInfo1Node)
        {
            var nodes = signatureInfo1Node.ChildNodes;

            _setUpId = nodes[0].InnerText ?? "";
            _signatureText = nodes[1].InnerText ?? "";
            _signatureImage = nodes[2].InnerText ?? "";
            _signatureComments = nodes[3].InnerText ?? "";
            _windowsVersion = nodes[4].InnerText ?? "";
            _officeVersion = nodes[5].InnerText ?? "";
            _applicationVersion = nodes[6].InnerText ?? "";
            _monitors = uint.Parse(nodes[7].InnerText ?? "");
            _horizontalResolution = uint.Parse(nodes[8].InnerText ?? "");
            _verticalResolution = uint.Parse(nodes[9].InnerText ?? "");
            _colorDepth = uint.Parse(nodes[10].InnerText ?? "");
            _signatureProviderID = nodes[11].InnerText ?? "";
            _signatureProviderUrl = nodes[12].InnerText ?? "";
            _signatureProviderDetails = int.Parse(nodes[13].InnerText ?? "-1");
            _signatureType = nodes[14].InnerText == "1" ? DigitalSignatureType.Invisible : DigitalSignatureType.SignatureLine;

            var delegateList1 = signatureInfo1Node.GetElementsByTagName("DelegateSuggestedSigner");
            if (delegateList1.Count != 0)
            {
                _delegateSuggestedSigner = ConvertUtil.ExcelDecodeString(delegateList1[0].InnerText);
            }

            var delegateList2 = signatureInfo1Node.GetElementsByTagName("DelegateSuggestedSigner2");
            if (delegateList2.Count != 0)
            {
                _delegateSuggestedSigner = ConvertUtil.ExcelDecodeString(delegateList2[0].InnerText);
            }

            var delegateSuggestedSignerEmailLst = signatureInfo1Node.GetElementsByTagName("DelegateSuggestedSignerEmail");
            if (delegateSuggestedSignerEmailLst.Count != 0)
            {
                _delegateSuggestedSigner = ConvertUtil.ExcelDecodeString(delegateSuggestedSignerEmailLst[0].InnerText);
            }

            var testNullVar = signatureInfo1Node.GetElementsByTagName("ManifestHashAlgorithm")[0];

            var manifestHashAlgorithmLst = signatureInfo1Node.GetElementsByTagName("ManifestHashAlgorithm");
            if (manifestHashAlgorithmLst.Count != 0)
            {
                string hashString = manifestHashAlgorithmLst[0].InnerText;
                _manifestHashAlgorithm = new Uri(hashString);
            }
        }

        //Required children
        internal string _setUpId = "";
        internal string _signatureText = "";
        //Base64 binary image string
        internal string _signatureImage;
        internal string _signatureComments;
        internal string _windowsVersion;
        internal string _officeVersion;
        internal string _applicationVersion;
        internal uint _monitors;
        internal uint _horizontalResolution;
        internal uint _verticalResolution;
        internal uint _colorDepth;
        internal string _signatureProviderID;
        internal string _signatureProviderUrl;
        internal int _signatureProviderDetails;
        internal DigitalSignatureType _signatureType;
        //Optional children
        internal string _delegateSuggestedSigner = null;
        internal string _delegateSuggestedSigner2 = null;
        internal string _delegateSuggestedSignerEmail = null;
        internal Uri _manifestHashAlgorithm = null;

        internal string GetXml()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append($"<SignatureInfoV1 xmlns=\"http://schemas.microsoft.com/office/2006/digsig\">");
            sb.Append($"<SetupID>{_setUpId}</SetupID>");
            sb.Append($"<SignatureText>{ConvertUtil.ExcelEscapeAndEncodeString(_signatureText)}</SignatureText>");
            sb.Append($"<SignatureImage>{_signatureImage}</SignatureImage>");
            sb.Append($"<SignatureComments>{ConvertUtil.ExcelEscapeAndEncodeString(_signatureComments)}</SignatureComments>");
            sb.Append($"<WindowsVersion>{_windowsVersion}</WindowsVersion>");
            sb.Append($"<OfficeVersion>{_officeVersion}</OfficeVersion>");
            sb.Append($"<ApplicationVersion>{_applicationVersion}</ApplicationVersion>");
            sb.Append($"<Monitors>{_monitors}</Monitors>");
            sb.Append($"<HorizontalResolution>{_horizontalResolution}</HorizontalResolution>");
            sb.Append($"<VerticalResolution>{_verticalResolution}</VerticalResolution>");
            sb.Append($"<ColorDepth>{_colorDepth}</ColorDepth>");
            sb.Append($"<SignatureProviderId>{_signatureProviderID}</SignatureProviderId>");
            sb.Append($"<SignatureProviderUrl>{_signatureProviderUrl}</SignatureProviderUrl>");
            sb.Append($"<SignatureProviderDetails>{_signatureProviderDetails}</SignatureProviderDetails>");
            sb.Append($"<SignatureType>{(int)_signatureType}</SignatureType>");

            if (_delegateSuggestedSigner != null)
            {
                sb.Append($"<DelegateSuggestedSigner>{ConvertUtil.ExcelEscapeAndEncodeString(_delegateSuggestedSigner)}</DelegateSuggestedSigner>");
            }

            if(_delegateSuggestedSigner2 != null)
            {
                sb.Append($"<DelegateSuggestedSigner2>{ConvertUtil.ExcelEscapeAndEncodeString(_delegateSuggestedSigner2)}</DelegateSuggestedSigner2>");
            }

            if (_delegateSuggestedSignerEmail != null)
            {
                sb.Append($"<DelegateSuggestedSignerEmail>{ConvertUtil.ExcelEscapeAndEncodeString(_delegateSuggestedSignerEmail)}</DelegateSuggestedSignerEmail>");
            }

            if(_manifestHashAlgorithm != null)
            {
                sb.Append($"<ManifestHashAlgorithm>{_manifestHashAlgorithm.AbsoluteUri}</ManifestHashAlgorithm>");
            }

            sb.Append($"</SignatureInfoV1>");

            return sb.ToString();
        }
    }
}
