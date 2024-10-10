using OfficeOpenXml.Packaging;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;


namespace OfficeOpenXml.DigitalSignatures
{
    internal class DigSigManifest
    {
        List<ManifestReference> manifestReferences;

        XmlDocument doc = new XmlDocument();

        internal void SortReferencesAndAddToDoc()
        {
            manifestReferences = manifestReferences.OrderBy(x => x._ref.Uri).ToList();
            foreach (var reference in manifestReferences)
            {
                ImportAndAddNode(reference.xmlDigSig);
            }
        }

        internal DigSigManifest()
        {
            manifestReferences = new List<ManifestReference>();
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
