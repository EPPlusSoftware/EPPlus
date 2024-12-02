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
        private Reference _ref;
        internal XmlElement xmlDigSig;
        private XmlDocument resultDoc = new XmlDocument();
        private string _uri;

        internal string SignatureMethod;
        internal string DigestMethod;
        internal PartType mRefType;

        internal string RefUri
        {
            get{ return _uri; }
        }

        public ManifestReference(PartWithXml aPart, string signatureMethod, string digestMethod)
        {
            SignatureMethod = signatureMethod;
            DigestMethod = digestMethod;

            mRefType = aPart.PartType;

            if (aPart.PartType == PartType.RelPart)
            {
                var relTransform = new RelTransform(aPart.Xml);
                CreateReference(aPart.UriKey, relTransform.GetOutputStream(), relTransform.TransformXml);
            }
            else
            {
                var xmlBytes = Encoding.UTF8.GetBytes(aPart.Xml);
                Stream xmlStream = RecyclableMemory.GetStream();
                xmlStream.Position = 0;
                xmlStream.Write(xmlBytes, 0, xmlBytes.Count());
                xmlStream.Position = 0;

                CreateReference(aPart.UriKey, xmlStream);
            }
        }

        public ManifestReference(XmlNode referenceNode)
        {
            resultDoc.LoadXml(referenceNode.OuterXml);
            xmlDigSig = (XmlElement)resultDoc.GetElementsByTagName("Reference")[0];
            _uri = xmlDigSig.GetAttribute("URI");
            var methodNode = xmlDigSig.GetElementsByTagName("DigestMethod")[0];
            DigestMethod = methodNode.Attributes.GetNamedItem("Algorithm").Value;
        }

        public ManifestReference(string uri, string xmlString, string signatureMethod, string digestMethod)
        {
            var xmlBytes = Encoding.UTF8.GetBytes(xmlString);

            Stream xmlStream = RecyclableMemory.GetStream();
            xmlStream.Position = 0;
            xmlStream.Write(xmlBytes, 0, xmlBytes.Count());
            xmlStream.Position = 0;

            SignatureMethod = signatureMethod;
            DigestMethod = digestMethod;

            CreateReference(uri, xmlStream);
        }

        public ManifestReference(string uri, Stream xmlStream, string signatureMethod, string digestMethod, string transformXml = null) 
        {
            xmlStream.Position = 0;
            SignatureMethod = signatureMethod;
            DigestMethod = digestMethod;
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
            signedXml.SignedInfo.SignatureMethod = SignatureMethod;

            _ref = new(doc);
            _ref.Uri = uri;
            _ref.DigestMethod = DigestMethod;

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
            }

            var element = (XmlElement)resultDoc.SelectSingleNode("//digSig:Reference", nsm);
            _uri = _ref.Uri;
            xmlDigSig = (XmlElement)resultDoc.GetElementsByTagName("Reference")[0];
        }
    }
}
