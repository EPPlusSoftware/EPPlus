using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.DigitalSignatures.XAdES
{
    internal class SignedDataObjectProperties
    {
        private string GetCommitmentTypeString(CommitmentType type)
        {
            switch (type)
            {
                case CommitmentType.None:
                    return "None";
                case CommitmentType.Approved:
                    return "Approved this document";
                case CommitmentType.Created:
                    return "Created this document";
                case CommitmentType.CreatedAndApproved:
                    return "Created and approved this document";
                default:
                    throw new NotImplementedException();
            }
        }

        private CommitmentType GetCommitmentTypeFromIdentifier(string identifier)
        {
            switch (identifier)
            {
                case "":
                    return CommitmentType.None;
                case "Approval":
                    return CommitmentType.Approved;
                case "Creation":
                    return CommitmentType.Created;
                case "Origin":
                    return CommitmentType.CreatedAndApproved;
                default:
                    throw new NotImplementedException();
            }
        }

        private string GetIdentifierFromCommitmentType(CommitmentType type)
        {
            switch (type)
            {
                case CommitmentType.None:
                    return "";
                case CommitmentType.Approved:
                    return "Approval";
                case CommitmentType.Created:
                    return "Creation";
                case CommitmentType.CreatedAndApproved:
                    return "Origin";
                default:
                    throw new NotImplementedException();
            }
        }

        const string origIdentifierForRemoval = "http://uri.etsi.org/01903/v1.2.2#ProofOf";
        const string origIdentifier = "http://uri.etsi.org/01903/v1.2.2#ProofOf{0}";
        string Identifier;
        string Description;
        List<string> CommitmentTypeQualifiers =  new List<string>();
        string Prefix = "xd";

        //Write
        internal SignedDataObjectProperties(string prefix, CommitmentType type, List<string> typeQualifiers)
        {
            Prefix = prefix;
            Description = GetCommitmentTypeString(type);

            if(type == CommitmentType.None)
            {
                Identifier = "";
            }
            else
            {
                Identifier = string.Format(origIdentifier, GetIdentifierFromCommitmentType(type));
            }

            CommitmentTypeQualifiers = typeQualifiers;
        }

        //Read
        internal SignedDataObjectProperties(string prefix, XmlElement SignedDataObjectPropertiesNode, List<string> typeQualifiers, ref CommitmentType type)
        {
            Prefix = prefix;
            var identifierNode = SignedDataObjectPropertiesNode.GetElementsByTagName($"{prefix}:Identifier")[0];
            if (identifierNode != null)
            {
                Identifier = identifierNode.InnerText;
                if (string.IsNullOrEmpty(Identifier))
                {
                    type = CommitmentType.None;
                }
                else
                {
                    var idForEnum = Identifier.Replace(origIdentifierForRemoval,"");
                    type = GetCommitmentTypeFromIdentifier(idForEnum);
                }
            }
            var descriptionNode = SignedDataObjectPropertiesNode.GetElementsByTagName($"{prefix}:Description")[0];
            if (descriptionNode != null)
            {
                Description = descriptionNode.InnerText;
                typeQualifiers.Add(Description);
            }

            var commitmentTypeQualifiers = SignedDataObjectPropertiesNode.GetElementsByTagName($"{prefix}:CommitmentTypeQualifier");
            for(int i = 0; i < commitmentTypeQualifiers.Count; i++)
            {
                CommitmentTypeQualifiers.Add(commitmentTypeQualifiers[i].InnerText);
                typeQualifiers.Add(commitmentTypeQualifiers[i].InnerText);
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
