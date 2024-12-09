using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.DigitalSignatures.XAdES
{
    internal class QualifyingProperties
    {
        internal SignedProperties SignedProps;
        string Prefix;
        string Target = "#idPackageSignature";
        string XadesNS = "http://uri.etsi.org/01903/v1.3.2#";

        internal QualifyingProperties(XmlElement signedPropertiesNode, AdditionalSignatureInfo info, List<string> TypeQualifiers, ref CommitmentType type)
        {
            var parentAttributes = signedPropertiesNode.ParentNode.Attributes;

            Prefix = signedPropertiesNode.Prefix;
            Target = parentAttributes.GetNamedItem("Target").InnerText;
            XadesNS = parentAttributes.GetNamedItem($"xmlns:{Prefix}").InnerText;

            SignedProps = new SignedProperties(signedPropertiesNode, info, TypeQualifiers, ref type);
        }

        internal QualifyingProperties(string prefix, X509Certificate2 cert, CommitmentType type, List<string> TypeQualifiers, AdditionalSignatureInfo info)
        {
            Prefix = prefix;
            SignedProps = new SignedProperties(cert, type, Prefix, TypeQualifiers, info);
        }

        internal string GetXML()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("<Object xmlns=\"http://www.w3.org/2000/09/xmldsig#\">");
            sb.Append($"<{Prefix}:QualifyingProperties xmlns:xd=\"{XadesNS}\" Target=\"{Target}\">");
            sb.Append(SignedProps.GetXML());
            sb.Append($"</{Prefix}:QualifyingProperties>");
            sb.Append("</Object>");
            return sb.ToString();
        }

        internal XmlDocument GetDocument() 
        {
            XmlDocument doc = new XmlDocument();

            doc.LoadXml(GetXML());

            var nsm = new XmlNamespaceManager(doc.NameTable);
            nsm.AddNamespace("xd", $"{XadesNS}");

            var node = doc.GetElementsByTagName("xd:SignedProperties")[0];

            doc.ImportNode(node, true);

            var element = doc.GetElementById("idSignedProperties");
            return doc;
        }
    }
}
