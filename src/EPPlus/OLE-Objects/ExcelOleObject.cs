using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.OleObject;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.CompundDocument;
using System;
using System.CodeDom;
using System.IO;
using System.Xml;

namespace OfficeOpenXml.OLE_Objects
{
    public class ExcelOleObject : XmlHelper, IDisposable
    {
        bool isLinkedObject = false;

        internal string progId;
        internal int shapeId;
        internal string relId; //relation to embedded object

        internal string prRelId; //relation to media

        internal CompoundDocument _document;
        internal ExcelWorksheet _worksheet;
        internal ExcelDrawing _drawing;
        internal ExcelOleObjects _oleObjects;

        public ExcelOleObject(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {
        }

        public ExcelOleObject(ExcelOleObjects oleObjects, ExcelWorksheet worksheet, XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            _worksheet = worksheet;
            _oleObjects = oleObjects;
            progId = GetXmlNodeString(topNode, "@progId");
            shapeId = int.Parse(GetXmlNodeString(topNode, "@shapeId"));
            relId = GetXmlNodeString(topNode, "@r:id");
            prRelId = GetXmlNodeString(topNode, "d:objectPr/@r:id");

            _drawing = _worksheet.Drawings.GetById(shapeId);

            LoadDocument();
        }

        internal void LoadDocument()
        {
            //TODO:
            //check if file is linked or embedded.


            var oleRel  = _worksheet.Part.GetRelationship(relId);
            var oleObj = UriHelper.ResolvePartUri(_worksheet.WorksheetUri, oleRel.TargetUri);
            var olePart = _worksheet._package.ZipPackage.GetPart(oleObj);
            var oleStream = (MemoryStream)olePart.GetStream(FileMode.Open, FileAccess.Read);
            _document = new CompoundDocument(oleStream);
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        internal static ExcelOleObject CreateOleObject(string filePath)
        {
            return null;
        }

        internal static ExcelOleObject GetOleObject(ExcelWorksheet worksheet, XmlNamespaceManager namespaceManager, ExcelOleObjects oleObjects, XmlNode node)
        {
            return new ExcelOleObject(oleObjects, worksheet, namespaceManager, node);
        }

        internal static ExcelDrawing GetOleDrawing(ExcelDrawings drawings, XmlElement drawNode, ExcelOleObject oleObject, ExcelGroupShape parent)
        {
            //create and return excel drawing
            return new OleObjectDrawing(oleObject, drawings, drawNode, null, null, parent);
        }



        //Ett Ole Objekt består av xml i worksheet och drawings med relationer till en embedding, printerSettings och Media
        //Embeddings har en CompoundDocument som innehåller filen samt en ole och CompObj filer. Word dokument ligger som .doc istället för som ett CompoundDocument
        //Microsoft_Word_Document, Microsoft_Word_Document1
        //printerSettings stödjer vi ej, men finns som en .bin fil.
        //Worksheet har 2 st relations id i sin nod. Den yttre har relation till embeddings objektet och den innre till dess Media.
        //Drawings har en xml nod som har relation till dess Media
        //Det finns även en VML i drawings
        //Media består av en bild på formatet .emf Både om man använder en ikon eller en bild av dokumentet.
        //Om filen länkas så skapas en externalLinks. Sökvägen till filen finns i _rels.
    }
}
