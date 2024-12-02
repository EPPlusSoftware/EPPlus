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

        internal DigSigManifest(ZipPackageXmlManifest preManifest, string signatureMethod, string digestMethod)
        {
            var root = doc.CreateElement("Manifest", "http://www.w3.org/2000/09/xmldsig#");
            doc.AppendChild(root);
            _signatureMethod = signatureMethod;
            _digestMethod = digestMethod;

            foreach(var part in preManifest.partXmlList)
            {
                var reference = new ManifestReference(part, signatureMethod, digestMethod);
                manifestReferences.Add(reference);
            }
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
        }

        internal DigSigManifest(string signatureMethod, string digestMethod)
        {
            var root = doc.CreateElement("Manifest", "http://www.w3.org/2000/09/xmldsig#");
            doc.AppendChild(root);
            _signatureMethod = signatureMethod;
            _digestMethod = digestMethod;
        }

        internal void AddRelsPartToManifest(string uri, string xmlString)
        {
            var relUri = uri + "?ContentType=" + ExcelPackage.schemaRelsExtension;

            RelTransform relTransform;

            relTransform = new RelTransform(xmlString);

            var manifestReference = new ManifestReference(relUri, relTransform.GetOutputStream(), _signatureMethod, _digestMethod, relTransform.TransformXml);
            manifestReferences.Add(manifestReference);
        }

        internal void AddPartToManifest(ZipPackagePart part, Stream xml)
        {
            var uri = part.Uri.OriginalString;
            var contentType = part.ContentType;
            var uriQuery = uri + "?ContentType=" + contentType;

            var manifestReference = new ManifestReference(uriQuery, xml, _signatureMethod, _digestMethod);
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
