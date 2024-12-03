using OfficeOpenXml.Packaging;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.DigitalSignatures
{
    internal class DigSigManifest
    {
        List<ManifestReference> manifestReferences = new();
        XmlDocument doc = new XmlDocument();

        private string _signatureMethod;
        private string _digestMethod;

        internal void SortReferencesAndAddToDoc()
        {
            manifestReferences = manifestReferences.OrderBy(x => x.RefUri).ToList();
            foreach (var reference in manifestReferences)
            {
                ImportAndAddNode(reference.xmlDigSig);
            }
        }

        internal DigSigManifest(ZipPackageXmlManifest preManifest, DigitalSignatureHashAlgorithm algorithm)
        {
            var root = doc.CreateElement("Manifest", "http://www.w3.org/2000/09/xmldsig#");
            doc.AppendChild(root);

            foreach(var part in preManifest.partXmlList)
            {
                var reference = new ManifestReference(part, algorithm);
                manifestReferences.Add(reference);
            }
            SortReferencesAndAddToDoc();
        }

        //Read manifest from signature
        internal DigSigManifest(XmlNode ManifestNode)
        {
            var signatureMethodNode = ManifestNode.OwnerDocument.DocumentElement.GetElementsByTagName("SignatureMethod")[0];
            _signatureMethod =  signatureMethodNode.Attributes.GetNamedItem("Algorithm").Value;
            doc.LoadXml(ManifestNode.OuterXml);
            var referenceElements = doc.GetElementsByTagName("Reference");
            foreach(XmlNode node in referenceElements)
            {
                var mReference = new ManifestReference(node);
                mReference.SignatureMethod = _signatureMethod;
                manifestReferences.Add(mReference);
            }
            _digestMethod = manifestReferences[0].DigestMethod;
            SortReferencesAndAddToDoc();
        }

        internal void ImportAndAddNode(XmlNode node)
        {
            var impNode = doc.ImportNode(node, true);
            doc.DocumentElement.AppendChild(impNode);
        }

        internal XmlDocument GetDoc()
        {
            return doc;
        }
    }
}
