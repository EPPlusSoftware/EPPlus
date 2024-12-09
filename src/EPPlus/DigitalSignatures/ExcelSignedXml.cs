using System.Security.Cryptography.Xml;
using System.Xml;

namespace OfficeOpenXml.DigitalSignatures
{
    internal class ExcelSignedXml : SignedXml
    {
        public ExcelSignedXml(XmlDocument document) : base(document)
        {

        }
        public ExcelSignedXml(XmlElement xmlElement)
        : base(xmlElement)
        {

        }

        public override XmlElement GetIdElement(XmlDocument document, string idValue)
        {
            XmlElement elem = base.GetIdElement(document, idValue);
            return elem;
        }
    }
}
