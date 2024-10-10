using System.Text;
using System.Xml;

namespace OfficeOpenXml.DigitalSignatures
{
    internal class SignatureInfoV2
    {
        AdditionalSignatureInfo _signatureInfo;

        public SignatureInfoV2(AdditionalSignatureInfo signatureInfo)
        {
            _signatureInfo = signatureInfo;
        }

        public SignatureInfoV2(XmlElement SignatureInfoV2Node, AdditionalSignatureInfo signatureInfo)
        {
            var nodes = SignatureInfoV2Node.ChildNodes;

            _signatureInfo = signatureInfo;

            var address1Lst = SignatureInfoV2Node.GetElementsByTagName("Address1");
            if (address1Lst.Count != 0)
            {
                _signatureInfo.Address1 = address1Lst[0].InnerText;
            }

            var address2Lst = SignatureInfoV2Node.GetElementsByTagName("Address2");
            if (address2Lst.Count != 0)
            {
                _signatureInfo.Address2 = address2Lst[0].InnerText;
            }
        }

        internal string GetXml()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append($"<SignatureInfoV2 xmlns=\"http://schemas.microsoft.com/office/2006/digsig\">");
            if(_signatureInfo.Address1 != null) 
            {
                sb.Append($"<Address1>{_signatureInfo.Address1}</Address1>");
            }
            if (_signatureInfo.Address2 != null)
            {
                sb.Append($"<Address2>{_signatureInfo.Address2}</Address2>");
            }
            sb.Append($"</SignatureInfoV2>");

            return sb.ToString();
        }
    }
}
