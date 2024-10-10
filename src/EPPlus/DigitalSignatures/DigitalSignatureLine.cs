using OfficeOpenXml.Drawing.OleObject;
using OfficeOpenXml.Drawing.Vml;
using System.IO;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.DigitalSignatures
{
    public class DigitalSignatureLine
    {
        ExcelDigitalSignature sig;
        ExcelVmlDrawingPictureCollection _drawings;
        ExcelWorksheet _ws;
        XmlNamespaceManager _nsm;
        ZipPackagePart part;
        const string relIdPath = "d:legacyDrawing/@r:id";

        SignatureLineEmf _emf;
        public ExcelVmlDrawingSignatureLine VmlDrawing;

        internal DigitalSignatureLine(ExcelWorksheet ws)
        {
            _ws = ws;
            _nsm = ws.NameSpaceManager;

            //Create Media
            int newID = 1;
            var Uri = XmlHelper.GetNewUri(ws._package.ZipPackage, "/xl/media/image{0}.emf", ref newID);
            part = ws._package.ZipPackage.CreatePart(Uri, "image/x-emf", CompressionLevel.None, "emf");
            //var rel = ws.Part.CreateRelationship(Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");

            part.SaveHandler = Save;

            VmlDrawing = ws.VmlDrawings.AddSignatureLine();

            ws.VmlDrawings.Part.CreateRelationship(UriHelper.GetRelativeUri(ws.VmlDrawings.Uri,Uri), TargetMode.Internal, ExcelPackage.schemaImage);
            _emf = VmlDrawing.Emf;
        }

        internal void Save(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
            //Init Zip
            stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
            stream.PutNextEntry(fileName);

            MemoryStream ms = (MemoryStream)part.GetStream(FileMode.Create, FileAccess.Write);
            _emf.SaveToStream(ms);

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

        internal void CreateSignatureLineDrawingPictureCollection()
        {
            var shapeTypeElement = (XmlElement)_drawings.VmlDrawingXml.DocumentElement.GetElementsByTagName("v:shapetype")[0];

            shapeTypeElement.SetAttribute("path", "m@4@5l@4@11@9@11@9@5xe");
            shapeTypeElement.SetAttribute("spt", oNameSpace, "75");
            shapeTypeElement.SetAttribute("id", "_x0000_t75");
            shapeTypeElement.SetAttribute("preferrelative", oNameSpace, "t");
            shapeTypeElement.SetAttribute("filled", "f");
            shapeTypeElement.SetAttribute("stroked", "f");

            var formulaElement = _drawings.VmlDrawingXml.CreateElement("v", "formulas", vNameSpace);
            shapeTypeElement.InsertAfter(formulaElement, shapeTypeElement.GetElementsByTagName("v:stroke")[0]);

            var doc = _drawings.VmlDrawingXml;
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

            var lockShapeTypeEl = doc.CreateElement("o", "lock", oNameSpace);
            shapeTypeElement.AppendChild(lockShapeTypeEl);
            lockShapeTypeEl.SetAttribute("ext", vNameSpace, "edit");
            lockShapeTypeEl.SetAttribute("aspectratio", "t");

            var pathElement = (XmlElement)shapeTypeElement.GetElementsByTagName("v:path")[0];
            pathElement.SetAttribute("extrusionok", oNameSpace, "f");
        }

        static void CreateAndSetFormulaElementOnNode(XmlElement formulaParentNode, XmlDocument document, string formula)
        {
            var f1 = document.CreateElement("v", "f", vNameSpace);
            f1.SetAttribute("eqn", formula);
            formulaParentNode.AppendChild(f1);
        }
    }
}
