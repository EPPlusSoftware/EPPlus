using System;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.DigitalSignatures
{
    /// <summary>
    /// Information about the signature including OSversion and office version
    /// </summary>
    public class SignatureInfoV1
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
                DelegateSuggestedSigner = delegateList1[0].InnerText;
            }

            var delegateList2 = SignatureInfo1Node.GetElementsByTagName("DelegateSuggestedSigner2");
            if (delegateList2.Count != 0)
            {
                DelegateSuggestedSigner = delegateList2[0].InnerText;
            }

            var DelegateSuggestedSignerEmailLst = SignatureInfo1Node.GetElementsByTagName("DelegateSuggestedSignerEmail");
            if (DelegateSuggestedSignerEmailLst.Count != 0)
            {
                DelegateSuggestedSigner = DelegateSuggestedSignerEmailLst[0].InnerText;
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
        public string SetUpId = "";
        public string SignatureText = "";
        //Base64 binary image string
        public string SignatureImage;
        public string SignatureComments;
        public string WindowsVersion;
        public string OfficeVersion;
        public string ApplicationVersion;
        public uint Monitors;
        public uint HorizontalResolution;
        public uint VerticalResolution;
        public uint ColorDepth;
        public string SignatureProviderID;
        public string SignatureProviderUrl;
        public int SignatureProviderDetails;
        public DigitalSignatureType SignatureType;
        //Optional children
        public string DelegateSuggestedSigner = null;
        public string DelegateSuggestedSigner2 = null;
        public string DelegateSuggestedSignerEmail = null;
        public Uri ManifestHashAlgorithm = null;

        internal string GetXml()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append($"<SignatureInfoV1 xmlns=\"http://schemas.microsoft.com/office/2006/digsig\">");
            sb.Append($"<SetupID>{SetUpId}</SetupID>");
            sb.Append($"<SignatureText>{SignatureText}</SignatureText>");
            sb.Append($"<SignatureImage>{SignatureImage}</SignatureImage>");
            sb.Append($"<SignatureComments>{SignatureComments}</SignatureComments>");
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
                sb.Append($"<DelegateSuggestedSigner>{DelegateSuggestedSigner}</DelegateSuggestedSigner>");
            }

            if(DelegateSuggestedSigner2 != null)
            {
                sb.Append($"<DelegateSuggestedSigner2>{DelegateSuggestedSigner2}</DelegateSuggestedSigner2>");
            }

            if (DelegateSuggestedSignerEmail != null)
            {
                sb.Append($"<DelegateSuggestedSignerEmail>{DelegateSuggestedSignerEmail}</DelegateSuggestedSignerEmail>");
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
