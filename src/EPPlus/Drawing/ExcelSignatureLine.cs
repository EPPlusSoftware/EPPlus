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
        ExcelWorksheet _ws;
        XmlNamespaceManager _nsm;
        ZipPackagePart part;
        const string relIdPath = "d:legacyDrawing/@r:id";

        internal ExcelSignatureLine(ExcelWorksheet ws, XmlNode topNode, XmlNamespaceManager ns, Guid lineId) : base(topNode, ns, lineId)
        {
            _ws = ws;
            _nsm = ws.NameSpaceManager;

            int newID = 1;
            var uri = GetNewUri(ws._package.ZipPackage, "/xl/media/image{0}.emf", ref newID);
            part = ws._package.ZipPackage.CreatePart(uri, "image/x-emf", CompressionLevel.None, "emf");
            part.SaveHandler = Save;

            ws.VmlDrawings.Part.CreateRelationship(UriHelper.GetRelativeUri(ws.VmlDrawings.Uri, uri), TargetMode.Internal, ExcelPackage.schemaImage);
        }

        internal void Save(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
            //Init Zip
            stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
            stream.PutNextEntry(fileName);

            MemoryStream ms = (MemoryStream)part.GetStream(FileMode.Create, FileAccess.Write);
            Emf.SaveToStream(ms);

            var b = (ms).ToArray();
            stream.Write(b, 0, b.Length);
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
