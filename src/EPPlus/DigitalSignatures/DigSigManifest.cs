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

        private void SortReferencesAndAddToDoc()
        {
            manifestReferences = manifestReferences.OrderBy(x => x.RefUri).ToList();
            foreach (var reference in manifestReferences)
            {
                ImportAndAddNode(reference.xmlDigSig);
            }
        }

        internal DigSigManifest(DigSigManifestContext preManifest, DigitalSignatureHashAlgorithm algorithm)
        {
            doc.PreserveWhitespace = true;
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
            doc.PreserveWhitespace = true;
            doc.LoadXml(ManifestNode.OuterXml);
            var referenceElements = doc.GetElementsByTagName("Reference");
            foreach(XmlNode node in referenceElements)
            {
                var mReference = new ManifestReference(node);
                manifestReferences.Add(mReference);
            }
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
