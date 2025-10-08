using OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering;
using OfficeOpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;


namespace OfficeOpenXml.CustomXml
{
    public class ExcelCustomXmlCollection
    {
        ExcelPackage _package;
        public ExcelCustomXmlCollection(ExcelPackage package)
        {
            _package = package;
        }
        public byte[]   ReadPQ()
        {
            var part = _package.ZipPackage.GetPart(new Uri("\\customXml\\item1.xml", UriKind.Relative));
            var xml = new XmlDocument();
            XmlHelper.LoadXmlSafe(xml, part.GetStream());

            NameTable nt = new NameTable();
            var ns = new XmlNamespaceManager(nt);
            ns.AddNamespace(string.Empty, "http://schemas.microsoft.com/DataMashup");

            var element = (XmlElement)xml.DocumentElement;
            return Convert.FromBase64String(element.InnerText);
        }
    }
}
