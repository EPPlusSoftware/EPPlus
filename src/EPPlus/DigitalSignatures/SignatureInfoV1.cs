using OfficeOpenXml.Utils;
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
            SignatureProviderID = defaultSignatureProvider;
            WindowsVersion = Environment.OSVersion.Version.ToString();

            if (eastAsianProvider)
            {
                SignatureProviderID = "{000CD6A4-0000-0000-C000-000000000046}";
            }
        }

        internal SignatureInfoV1(XmlElement SignatureInfo1Node)
        {
            var nodes = SignatureInfo1Node.ChildNodes;

            SetUpId = nodes[0].InnerText ?? "";
            SignatureText = nodes[1].InnerText ?? "";
            SignatureImage = nodes[2].InnerText ?? "";
            SignatureComments = nodes[3].InnerText ?? "";
            WindowsVersion = nodes[4].InnerText ?? "";
            OfficeVersion = nodes[5].InnerText ?? "";
            ApplicationVersion = nodes[6].InnerText ?? "";
            Monitors = uint.Parse(nodes[7].InnerText ?? "");
            HorizontalResolution = uint.Parse(nodes[8].InnerText ?? "");
            VerticalResolution = uint.Parse(nodes[9].InnerText ?? "");
            ColorDepth = uint.Parse(nodes[10].InnerText ?? "");
            SignatureProviderID = nodes[11].InnerText ?? "";
            SignatureProviderUrl = nodes[12].InnerText ?? "";
            SignatureProviderDetails = int.Parse(nodes[13].InnerText ?? "-1");
            SignatureType = nodes[14].InnerText == "1" ? DigitalSignatureType.Invisible : DigitalSignatureType.SignatureLine;

            var delegateList1 = SignatureInfo1Node.GetElementsByTagName("DelegateSuggestedSigner");
            if (delegateList1.Count != 0)
            {
                DelegateSuggestedSigner = ConvertUtil.ExcelDecodeString(delegateList1[0].InnerText);
            }

            var delegateList2 = SignatureInfo1Node.GetElementsByTagName("DelegateSuggestedSigner2");
            if (delegateList2.Count != 0)
            {
                DelegateSuggestedSigner = ConvertUtil.ExcelDecodeString(delegateList2[0].InnerText);
            }

            var DelegateSuggestedSignerEmailLst = SignatureInfo1Node.GetElementsByTagName("DelegateSuggestedSignerEmail");
            if (DelegateSuggestedSignerEmailLst.Count != 0)
            {
                DelegateSuggestedSigner = ConvertUtil.ExcelDecodeString(DelegateSuggestedSignerEmailLst[0].InnerText);
            }

            var testNullVar = SignatureInfo1Node.GetElementsByTagName("ManifestHashAlgorithm")[0];

            var ManifestHashAlgorithmLst = SignatureInfo1Node.GetElementsByTagName("ManifestHashAlgorithm");
            if (ManifestHashAlgorithmLst.Count != 0)
            {
                string hashString = ManifestHashAlgorithmLst[0].InnerText;
                ManifestHashAlgorithm = new Uri(hashString);
            }
        }

        //Required children
        internal string SetUpId = "";
        internal string SignatureText = "";
        //Base64 binary image string
        internal string SignatureImage;
        internal string SignatureComments;
        internal string WindowsVersion;
        internal string OfficeVersion;
        internal string ApplicationVersion;
        internal uint Monitors;
        internal uint HorizontalResolution;
        internal uint VerticalResolution;
        internal uint ColorDepth;
        internal string SignatureProviderID;
        internal string SignatureProviderUrl;
        internal int SignatureProviderDetails;
        internal DigitalSignatureType SignatureType;
        //Optional children
        internal string DelegateSuggestedSigner = null;
        internal string DelegateSuggestedSigner2 = null;
        internal string DelegateSuggestedSignerEmail = null;
        internal Uri ManifestHashAlgorithm = null;

        internal string GetXml()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append($"<SignatureInfoV1 xmlns=\"http://schemas.microsoft.com/office/2006/digsig\">");
            sb.Append($"<SetupID>{SetUpId}</SetupID>");
            sb.Append($"<SignatureText>{ConvertUtil.ExcelEscapeAndEncodeString(SignatureText)}</SignatureText>");
            sb.Append($"<SignatureImage>{SignatureImage}</SignatureImage>");
            sb.Append($"<SignatureComments>{ConvertUtil.ExcelEscapeAndEncodeString(SignatureComments)}</SignatureComments>");
            sb.Append($"<WindowsVersion>{WindowsVersion}</WindowsVersion>");
            sb.Append($"<OfficeVersion>{OfficeVersion}</OfficeVersion>");
            sb.Append($"<ApplicationVersion>{ApplicationVersion}</ApplicationVersion>");
            sb.Append($"<Monitors>{Monitors}</Monitors>");
            sb.Append($"<HorizontalResolution>{HorizontalResolution}</HorizontalResolution>");
            sb.Append($"<VerticalResolution>{VerticalResolution}</VerticalResolution>");
            sb.Append($"<ColorDepth>{ColorDepth}</ColorDepth>");
            sb.Append($"<SignatureProviderId>{SignatureProviderID}</SignatureProviderId>");
            sb.Append($"<SignatureProviderUrl>{SignatureProviderUrl}</SignatureProviderUrl>");
            sb.Append($"<SignatureProviderDetails>{SignatureProviderDetails}</SignatureProviderDetails>");
            sb.Append($"<SignatureType>{(int)SignatureType}</SignatureType>");

            if (DelegateSuggestedSigner != null)
            {
                sb.Append($"<DelegateSuggestedSigner>{ConvertUtil.ExcelEscapeAndEncodeString(DelegateSuggestedSigner)}</DelegateSuggestedSigner>");
            }

            if(DelegateSuggestedSigner2 != null)
            {
                sb.Append($"<DelegateSuggestedSigner2>{ConvertUtil.ExcelEscapeAndEncodeString(DelegateSuggestedSigner2)}</DelegateSuggestedSigner2>");
            }

            if (DelegateSuggestedSignerEmail != null)
            {
                sb.Append($"<DelegateSuggestedSignerEmail>{ConvertUtil.ExcelEscapeAndEncodeString(DelegateSuggestedSignerEmail)}</DelegateSuggestedSignerEmail>");
            }

            if(ManifestHashAlgorithm != null)
            {
                sb.Append($"<ManifestHashAlgorithm>{ManifestHashAlgorithm.AbsoluteUri}</ManifestHashAlgorithm>");
            }

            sb.Append($"</SignatureInfoV1>");

            return sb.ToString();
        }
    }
}
