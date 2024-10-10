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

        internal SignedSignatureProperites SignatureProps;
        SignedDataObjectProperties DataObjectProps;

        internal SignedProperties(XmlElement signedPropertiesNode, AdditionalSignatureInfo info)
        {
            Prefix = signedPropertiesNode.Prefix;
            if(signedPropertiesNode.ChildNodes.Count > 0)
            {
                SignatureProps = new SignedSignatureProperites(Prefix, (XmlElement)signedPropertiesNode.ChildNodes[0], info);
            }
            if(signedPropertiesNode.ChildNodes.Count > 1)
            {
                DataObjectProps = new SignedDataObjectProperties(Prefix, (XmlElement)signedPropertiesNode.ChildNodes[1]);
            }
        }


        internal SignedProperties(X509Certificate2 cert, string desc, string prefix, List<string> TypeQualifiers, AdditionalSignatureInfo info)
        {
            Prefix = prefix;
            SignatureProps = new SignedSignatureProperites(Prefix, cert, info);
            DataObjectProps = new SignedDataObjectProperties(Prefix, desc, TypeQualifiers);
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
