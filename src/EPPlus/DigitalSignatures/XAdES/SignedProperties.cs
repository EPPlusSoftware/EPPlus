using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.DigitalSignatures.XAdES
{
    internal class SignedProperties
    {
        string Id = "idSignedProperties";
        string Prefix = "xd";

        internal SignedSignatureProperties SignatureProps;
        internal SignedDataObjectProperties DataObjectProps;

        internal SignedProperties(XmlElement signedPropertiesNode, AdditionalSignatureInfo info, List<string> TypeQualifiers, ref CommitmentType type)
        {
            Prefix = signedPropertiesNode.Prefix;
            if(signedPropertiesNode.ChildNodes.Count > 0)
            {
                SignatureProps = new SignedSignatureProperties(Prefix, (XmlElement)signedPropertiesNode.ChildNodes[0], info);
            }
            if(signedPropertiesNode.ChildNodes.Count > 1)
            {
                DataObjectProps = new SignedDataObjectProperties(Prefix, (XmlElement)signedPropertiesNode.ChildNodes[1], TypeQualifiers, ref type);
            }
        }


        internal SignedProperties(X509Certificate2 cert, CommitmentType type, string prefix, List<string> TypeQualifiers, AdditionalSignatureInfo info)
        {
            Prefix = prefix;
            SignatureProps = new SignedSignatureProperties(Prefix, cert, info);
            DataObjectProps = new SignedDataObjectProperties(Prefix, type, TypeQualifiers);
        }

        internal string GetXML()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append($"<{Prefix}:SignedProperties Id=\"{Id}\">");
            sb.Append(SignatureProps.GetXML());
            sb.Append(DataObjectProps.GetXML());
            sb.Append($"</{Prefix}:SignedProperties>");

            return sb.ToString();
        }
    }
}
