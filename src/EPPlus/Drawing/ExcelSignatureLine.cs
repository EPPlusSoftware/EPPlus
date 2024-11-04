using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System.IO;
using System.Xml;
using OfficeOpenXml.Drawing.Vml;
using System;

namespace OfficeOpenXml.Drawing
{
    public class ExcelSignatureLine : ExcelVmlDrawingSignatureLine
    {
        ExcelDigitalSignature sig;
        ExcelVmlDrawingPictureCollection _drawings;
        internal ExcelWorksheet Worksheet;
        XmlNamespaceManager _nsm;
        //private ZipPackagePart part;
        const string relIdPath = "d:legacyDrawing/@r:id";

        public string SignatureText = "";
        public byte[] SignatureImage;

        internal ExcelSignatureLine(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns, Guid lineId) : base(topNode, ns, lineId)
        {
            Worksheet = ws;
            _nsm = ws.NameSpaceManager;
        }

        internal ExcelSignatureLine(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns) : base(topNode, ns)
        {
            Worksheet = ws;
            _nsm = ws.NameSpaceManager;
        }

        const string vNameSpace = "urn:schemas-microsoft-com:vml";
        const string oNameSpace = "urn:schemas-microsoft-com:office:office";
        const string xNameSpace = "urn:schemas-microsoft-com:office:excel";

        internal static void CreateFormulaElementAsChildOf(XmlNode node)
        {
            var doc = node.OwnerDocument;

            var formulaElement = doc.CreateElement("v", "formulas", vNameSpace);
            node.AppendChild(formulaElement);

            CreateAndSetFormulaElementOnNode(formulaElement, doc, "if lineDrawn pixelLineWidth 0");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "sum @0 1 0");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "sum 0 0 @1");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "prod @2 1 2");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "prod @3 21600 pixelWidth");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "prod @3 21600 pixelHeight");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "sum @0 0 1");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "prod @6 1 2");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "prod @7 21600 pixelWidth");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "sum @8 21600 0");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "prod @7 21600 pixelHeight");
            CreateAndSetFormulaElementOnNode(formulaElement, doc, "sum @10 21600 0");
        }

        static void CreateAndSetFormulaElementOnNode(XmlElement formulaParentNode, XmlDocument document, string formula)
        {
            var f1 = document.CreateElement("v", "f", vNameSpace);
            f1.SetAttribute("eqn", formula);
            formulaParentNode.AppendChild(f1);
        }
    }
}
