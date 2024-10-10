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

        internal QualifyingProperties(XmlElement signedPropertiesNode, AdditionalSignatureInfo info)
        {
            var parentAttributes = signedPropertiesNode.ParentNode.Attributes;

            Prefix = signedPropertiesNode.Prefix;
            Target = parentAttributes.GetNamedItem("Target").InnerText;
            XadesNS = parentAttributes.GetNamedItem($"xmlns:{Prefix}").InnerText;

            SignedProps = new SignedProperties(signedPropertiesNode, info);
        }

        internal QualifyingProperties(string prefix, X509Certificate2 cert, string decription, List<string> TypeQualifiers, AdditionalSignatureInfo info)
        {
            Prefix = prefix;
            SignedProps = new SignedProperties(cert, decription, Prefix, TypeQualifiers, info);
        }

        internal string GetXML()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("<Object xmlns=\"http://www.w3.org/2000/09/xmldsig#\">");
            sb.Append($"<{Prefix}:QualifyingProperties xmlns:xd=\"{XadesNS}\" Target=\"{Target}\">");
            sb.Append(SignedProps.GetXML());
            //sb.Append(unsignedProperties.GetXML());
            sb.Append($"</{Prefix}:QualifyingProperties>");
            sb.Append("</Object>");
            return sb.ToString();
        }

        internal XmlDocument GetDocument() 
        {
            XmlDocument doc = new XmlDocument();
            //doc.CreateElement("xd", "QualifyingProperties", $"{XadesNS}");

            doc.LoadXml(GetXML());

            var nsm = new XmlNamespaceManager(doc.NameTable);
            nsm.AddNamespace("xd", $"{XadesNS}");

            var node = doc.GetElementsByTagName("xd:SignedProperties")[0];

            doc.ImportNode(node, true);

            var element = doc.GetElementById("idSignedProperties");

            //XmlDocument objectDocument = new XmlDocument();

            ////var nsm2 = new XmlNamespaceManager(objectDocument.NameTable);
            ////nsm2.AddNamespace("xd", "XadesNS");

            //var rootObject = objectDocument.CreateElement("Object", "http://www.w3.org/2000/09/xmldsig#");
            ////objectDocument.AppendChild(rootObject);

            //foreach(XmlNode node in doc.DocumentElement.ChildNodes)
            //{
            //    rootObject.AppendChild(objectDocument.ImportNode(node, true));
            //}

            //var importedNode = objectDocument.ImportNode(doc.DocumentElement, true);
            //objectDocument.AppendChild(importedNode);
            return doc;
        }

        //SignedDataObjectProperties signedDataObjectProperties;
        //UnsignedProperties unsignedProperties;
        //UnsignedSignatureProperties unsignedSignatures;
    }
}
