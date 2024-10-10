using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.DigitalSignatures.XAdES
{
    internal class SignedDataObjectProperties
    {
        string Identifier = "http://uri.etsi.org/01903/v1.2.2#ProofOfOrigin";
        string Description;
        List<string> CommitmentTypeQualifiers;
        string Prefix = "xd";

        internal SignedDataObjectProperties(string prefix, string desc, List<string> typeQualifiers)
        {
            Prefix = prefix;
            Description = desc;
            CommitmentTypeQualifiers = typeQualifiers;
        }

        internal SignedDataObjectProperties(string prefix, XmlElement SignedDataObjectPropertiesNode)
        {
            Prefix = prefix;
            var identifierNode = SignedDataObjectPropertiesNode.GetElementsByTagName($"{prefix}:Identifier")[0];
            if (identifierNode != null)
            {
                Identifier = identifierNode.InnerText;
            }
            var descriptionNode = SignedDataObjectPropertiesNode.GetElementsByTagName($"{prefix}:Description")[0];
            if (descriptionNode != null)
            {
                Description = descriptionNode.InnerText;
            }
        }

        internal string GetXML()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append($"<{Prefix}:SignedDataObjectProperties>");
            sb.Append($"<{Prefix}:CommitmentTypeIndication>");

            sb.Append($"<{Prefix}:CommitmentTypeId>");

            sb.Append($"<{Prefix}:Identifier>{Identifier}</{Prefix}:Identifier>");
            sb.Append($"<{Prefix}:Description>{Description}</{Prefix}:Description>");

            sb.Append($"</{Prefix}:CommitmentTypeId>");

            sb.Append($"<{Prefix}:AllSignedDataObjects></{Prefix}:AllSignedDataObjects>");

            sb.Append($"<{Prefix}:CommitmentTypeQualifiers>");
            //Purposes for signing document
            for(int i = 0; i< CommitmentTypeQualifiers.Count; i++)
            {
                sb.Append($"<{Prefix}:CommitmentTypeQualifier>{CommitmentTypeQualifiers[i]}</{Prefix}:CommitmentTypeQualifier>");
            }
            sb.Append($"</{Prefix}:CommitmentTypeQualifiers>");

            sb.Append($"</{Prefix}:CommitmentTypeIndication>");
            sb.Append($"</{Prefix}:SignedDataObjectProperties>");

            return sb.ToString();
        }
    }
}
