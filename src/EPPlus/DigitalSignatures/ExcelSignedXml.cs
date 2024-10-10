using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeOpenXml.DigitalSignatures
{
    public class ExcelSignedXml : SignedXml
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

            //if(elem == null && document != null) 
            //{
            //    var nodes = document.SelectNodes("//*[@Id]");
            //    foreach(XmlNode node in nodes) 
            //    {
            //        if(node.Attributes.GetNamedItem("Id").Value == idValue)
            //        {
            //            return elem;
            //        }
            //    }
            //}

            return elem;
        }
    }
}
