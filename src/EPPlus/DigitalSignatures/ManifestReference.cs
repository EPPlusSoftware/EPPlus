using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Xml;
using System;
using OfficeOpenXml.Utils.EncodingUtils;

namespace OfficeOpenXml.DigitalSignatures
{
    internal class ManifestReference
    {
        //private Reference _ref;
        internal XmlElement xmlDigSig;
        internal XmlElement xmlDigSigAlt;

        private XmlDocument resultDoc = new XmlDocument();
        private string _uri;
        const string TemplateXml = "<Root xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><Reference URI=\"{0}\">{1}<DigestMethod Algorithm=\"{2}\"/><DigestValue>{3}</DigestValue></Reference></Root>";
        const string TemplateTransforms = "<Transforms>{0}<Transform Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\"/></Transforms>";

        internal string DigestMethod;
        internal ePartType mRefType;

        internal string RefUri
        {
            get{ return _uri; }
        }

        public ManifestReference(PartWithXml xmlPart, DigitalSignatureHashAlgorithm algorithm)
        {
            mRefType = xmlPart.PartType;
            _uri = xmlPart.UriKey;

            if (xmlPart.PartType == ePartType.RelPart)
            {
                var relTransform = new RelTransform(xmlPart.StringData);
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
                    var resString = string.Format(TemplateXml, xmlPart.UriKey, "", DigestMethods.GetDigestMethod(algorithm), xmlPart.StringData);
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
    }
}
