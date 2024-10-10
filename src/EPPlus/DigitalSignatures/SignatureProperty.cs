using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.DigitalSignatures
{
    public class SignatureProperty
    {
        internal string Target;
        internal string Id;
        internal SignatureTime PropertySignatureTime = null;
        XmlDocument doc;
        SignatureInfoV1 sigInfo1;
        SignatureInfoV2 sigInfo2;

        internal SignatureProperty(XmlElement signaturePropertyElement, AdditionalSignatureInfo signatureInfo) 
        {
            var infoNodes = signaturePropertyElement.ChildNodes;
            sigInfo1 = new SignatureInfoV1((XmlElement)infoNodes[0]);
            if(infoNodes.Count > 1)
            {
                sigInfo2 = new SignatureInfoV2((XmlElement)infoNodes[1], signatureInfo);
            }
        }

        internal SignatureProperty(string target, string id) 
        {
            Target = target;
            Id = id;

            doc = new XmlDocument() { PreserveWhitespace = true };
            var root = doc.CreateElement("SignatureProperties", "http://www.w3.org/2000/09/xmldsig#");
            doc.AppendChild(root);
        }

        internal SignatureProperty(string target, string id, DateTime timeValue)
        {
            Target = target;
            Id = id;
            PropertySignatureTime = new SignatureTime(timeValue);

            doc = new XmlDocument();
            var root = doc.CreateElement("SignatureProperties", "http://www.w3.org/2000/09/xmldsig#");
            doc.AppendChild(root);
        }

        internal void CreateSignatureInfo(AdditionalSignatureInfo signatureInfo)
        {
            sigInfo1 = new SignatureInfoV1();
            if(signatureInfo.Address1 != null || signatureInfo.Address2 != null) 
            {
                sigInfo2 = new SignatureInfoV2(signatureInfo);
            }
        }

        internal string GetXML()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append($"<SignatureProperties xmlns=\"http://www.w3.org/2000/09/xmldsig#\">");

            sb.Append($"<SignatureProperty  Id=\"{Id}\" Target=\"{Target}\">");
            if (sigInfo1 != null)
            { 
                sb.Append($"{sigInfo1.GetXml()}");
            }
            if (sigInfo2 != null)
            {
                sb.Append($"{sigInfo2.GetXml()}");
            }
            sb.Append($"</SignatureProperty>");

            if (PropertySignatureTime != null)
            {
                sb.Append($"<SignatureProperty Id=\"{Id}\" Target=\"{Target}\">");
                sb.Append($"{PropertySignatureTime.GetXml()}");
                sb.Append($"</SignatureProperty>");
            }
            sb.Append($"</SignatureProperties>");

            return sb.ToString();
        }

        internal XmlDocument GetXMLDocument()
        {
            var tempDoc = new XmlDocument() { PreserveWhitespace = true };
            var xmlStuff = GetXML();
            tempDoc.LoadXml(xmlStuff);
            var node = doc.ImportNode(tempDoc.DocumentElement, true);
            doc.DocumentElement.AppendChild(node.ChildNodes[0]);
            return doc;
        }

    }
}
