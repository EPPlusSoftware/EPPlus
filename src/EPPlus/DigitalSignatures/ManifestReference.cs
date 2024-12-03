using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Xml;
using OfficeOpenXml.Utils;
using System;

namespace OfficeOpenXml.DigitalSignatures
{
    internal class ManifestReference
    {
        private Reference _ref;
        internal XmlElement xmlDigSig;
        internal XmlElement xmlDigSigAlt;

        private XmlDocument resultDoc = new XmlDocument();
        private string _uri;
        const string TemplateXml = "<Root xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><Reference URI=\"{0}\">{1}<DigestMethod Algorithm=\"{2}\"/><DigestValue>{3}</DigestValue></Reference></Root>";
        const string TemplateTransforms = "<Transforms>{0}<Transform Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\"/></Transforms>";

        internal string SignatureMethod;
        internal string DigestMethod;
        internal ePartType mRefType;

        internal string RefUri
        {
            get{ return _uri; }
        }

        public ManifestReference(PartWithXml xmlPart, DigitalSignatureHashAlgorithm algorithm)
        {
            mRefType = xmlPart.PartType;

            if (xmlPart.PartType == ePartType.RelPart)
            {
                var relTransform = new RelTransform(Encoding.UTF8.GetString(xmlPart.Bytes));
                var transform = new XmlDsigC14NTransform();
                var doc = new XmlDocument();
                doc.LoadXml(relTransform.GetOutputXML());
                transform.LoadInput(doc);

                MemoryStream ms = (MemoryStream)transform.GetOutput();
                var digestValue = EncodeUtil.HashAndEncodeBytes(ms.ToArray(), algorithm);

                var transformStr = string.Format(TemplateTransforms, relTransform.TransformXml);
                var resString = string.Format(TemplateXml, xmlPart.UriKey, transformStr, DigestMethods.GetDigestMethod(algorithm), digestValue);
                resultDoc.LoadXml(resString);
                xmlDigSig = (XmlElement)resultDoc.GetElementsByTagName("Reference")[0];
            }
            else
            {
                try
                {
                    var digestValue = EncodeUtil.HashAndEncodeBytes(xmlPart.Bytes, algorithm);
                    var resString = string.Format(TemplateXml, xmlPart.UriKey, "", DigestMethods.GetDigestMethod(algorithm), digestValue);
                    resultDoc.LoadXml(resString);
                    xmlDigSig = (XmlElement)resultDoc.GetElementsByTagName("Reference")[0];
                }
                catch(Exception e)
                {
                    throw new InvalidOperationException($"InnerException:{e}, message: {e.Message}");
                }
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

        //public ManifestReference(string uri, string xmlString, string signatureMethod, string digestMethod)
        //{
        //    var xmlBytes = Encoding.UTF8.GetBytes(xmlString);

        //    Stream xmlStream = RecyclableMemory.GetStream();
        //    xmlStream.Position = 0;
        //    xmlStream.Write(xmlBytes, 0, xmlBytes.Count());
        //    xmlStream.Position = 0;

        //    SignatureMethod = signatureMethod;
        //    DigestMethod = digestMethod;

        //    CreateReference(uri, xmlStream);
        //}

        //public ManifestReference(string uri, Stream xmlStream, string signatureMethod, string digestMethod, string transformXml = null) 
        //{
        //    xmlStream.Position = 0;
        //    SignatureMethod = signatureMethod;
        //    DigestMethod = digestMethod;
        //    CreateReference(uri, xmlStream, transformXml);
        //}

        //void CreateReference(string uri, Stream doc, string transformXml = null)
        //{
        //    RSACryptoServiceProvider rsaKey = new();

        //    SignedXml signedXml = new()
        //    {
        //        SigningKey = rsaKey,
        //    };

        //    signedXml.SignedInfo.CanonicalizationMethod = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315";
        //    signedXml.SignedInfo.SignatureMethod = SignatureMethod;

        //    _ref = new(doc);
        //    _ref.Uri = uri;
        //    _ref.DigestMethod = DigestMethod;

        //    if (transformXml != null)
        //    {
        //        _ref.AddTransform(new XmlDsigC14NTransform());
        //    }

        //    signedXml.AddReference(_ref);
        //    signedXml.ComputeSignature();

        //    var retXml = signedXml.GetXml();

        //    resultDoc.LoadXml(retXml.OuterXml);

        //    var nsm = new XmlNamespaceManager(resultDoc.NameTable);
        //    nsm.AddNamespace("digSig", retXml.NamespaceURI);

        //    if (transformXml != null)
        //    {
        //        var transforms = resultDoc.SelectSingleNode(".//digSig:Transforms", nsm);
        //        transforms.InnerXml = transformXml + transforms.InnerXml;
        //    }

        //    var element = (XmlElement)resultDoc.SelectSingleNode("//digSig:Reference", nsm);
        //    _uri = _ref.Uri;
        //    xmlDigSig = (XmlElement)resultDoc.GetElementsByTagName("Reference")[0];
        //}
    }
}
