using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.DigitalSignatures
{
    internal class ManifestReference
    {
        internal Reference _ref;
        internal XmlElement xmlDigSig;
        internal string xmlString;
        internal XmlDocument resultDoc = new XmlDocument();
        internal XmlNodeList children;

        public ManifestReference(string uri, string xmlString)
        {
            var xmlBytes = Encoding.UTF8.GetBytes(xmlString);

            Stream xmlStream = RecyclableMemory.GetStream();
            xmlStream.Position = 0;
            xmlStream.Write(xmlBytes, 0, xmlBytes.Count());
            xmlStream.Position = 0;

            CreateReference(uri, xmlStream);
        }

        public ManifestReference(string uri, Stream xmlStream, string transformXml = null) 
        {
            xmlStream.Position = 0;
            CreateReference(uri, xmlStream, transformXml);
        }

        void CreateReference(string uri, Stream doc, string transformXml = null)
        {
            RSACryptoServiceProvider rsaKey = new();

            SignedXml signedXml = new()
            {
                SigningKey = rsaKey,
            };

            signedXml.SignedInfo.CanonicalizationMethod = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315";
            signedXml.SignedInfo.SignatureMethod = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";

            _ref = new(doc);
            _ref.Uri = uri;
            _ref.DigestMethod = "http://www.w3.org/2000/09/xmldsig#sha1";

            if (transformXml != null)
            {
                _ref.AddTransform(new XmlDsigC14NTransform());
            }

            signedXml.AddReference(_ref);
            signedXml.ComputeSignature();

            var retXml = signedXml.GetXml();

            resultDoc.LoadXml(retXml.OuterXml);

            var nsm = new XmlNamespaceManager(resultDoc.NameTable);
            nsm.AddNamespace("digSig", retXml.NamespaceURI);

            if (transformXml != null)
            {
                var transforms = resultDoc.SelectSingleNode(".//digSig:Transforms", nsm);

                transforms.InnerXml = transformXml + transforms.InnerXml;
                var text = transforms.InnerText;
            }

            var element = (XmlElement)resultDoc.SelectSingleNode("//digSig:Reference", nsm);
            children = element.ParentNode.ChildNodes;
            xmlDigSig = (XmlElement)resultDoc.GetElementsByTagName("Reference")[0];
        }

    }
}
