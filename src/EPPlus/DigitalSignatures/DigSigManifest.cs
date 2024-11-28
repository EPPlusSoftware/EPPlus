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

        internal void SortReferencesAndAddToDoc()
        {
            manifestReferences = manifestReferences.OrderBy(x => x.RefUri).ToList();
            foreach (var reference in manifestReferences)
            {
                ImportAndAddNode(reference.xmlDigSig);
            }
        }

        //Read manifest from signature
        internal DigSigManifest(XmlNode ManifestNode)
        {
            doc.LoadXml(ManifestNode.OuterXml);
            var referenceElements = doc.GetElementsByTagName("Reference");
            foreach(XmlNode node in referenceElements)
            {
                var mReference = new ManifestReference(node);
                manifestReferences.Add(mReference);
            }
        }

        internal DigSigManifest()
        {
            var root = doc.CreateElement("Manifest", "http://www.w3.org/2000/09/xmldsig#");
            doc.AppendChild(root);
        }

        internal void AddRelsPartToManifest(string uri, string xmlString)
        {
            var relUri = uri + "?ContentType=" + ExcelPackage.schemaRelsExtension;

            RelTransform relTransform;

            relTransform = new RelTransform(xmlString);

            var manifestReference = new ManifestReference(relUri, relTransform.GetOutputStream(), relTransform.TransformXml);
            manifestReferences.Add(manifestReference);
        }

        internal void AddPartToManifest(ZipPackagePart part, Stream xml)
        {
            var uri = part.Uri.OriginalString;
            var contentType = part.ContentType;
            var uriQuery = uri + "?ContentType=" + contentType;

            var manifestReference = new ManifestReference(uriQuery, xml);
            manifestReferences.Add(manifestReference);
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
