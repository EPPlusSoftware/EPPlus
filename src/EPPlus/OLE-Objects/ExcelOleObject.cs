using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.OleObject;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Utils.CompundDocument;
using System;
using System.CodeDom;
using System.IO;
using System.Reflection.Emit;
using System.Xml;

namespace OfficeOpenXml.OLE_Objects
{
    public class ExcelOleObject : XmlHelper, IDisposable
    {
        public bool isExternalLink = false;

        internal string progId;
        internal int shapeId;  //relation to drawing
        internal string relId; //relation to embedded object
        internal string link;
        internal string prRelId; //relation to media

        internal CompoundDocument _document;
        internal ExcelExternalOleLink _externalLink;

        internal ExcelWorksheet _worksheet;
        internal ExcelDrawing _drawing;

        internal ExcelOleObjects _oleObjects;

        public ExcelOleObject(ExcelOleObjects oleObjects, ExcelWorksheet worksheet, XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
            _worksheet = worksheet;
            _oleObjects = oleObjects;
            progId = GetXmlNodeString(topNode, "@progId");
            shapeId = int.Parse(GetXmlNodeString(topNode, "@shapeId"));
            relId = GetXmlNodeString(topNode, "@r:id");
            link = GetXmlNodeString(topNode, "@link");
            prRelId = GetXmlNodeString(topNode, "d:objectPr/@r:id");

            //check if object is a linked object
            if(string.IsNullOrEmpty(link))
            {
                isExternalLink = false;
                LoadDocument();
            }
            else
            {
                isExternalLink = true;
                LoadExternalLink();
            }
        }

        public ExcelDrawing Drawing
        {
            get
            {
                if (_drawing == null)
                {
                    _drawing = _worksheet.Drawings.GetById(shapeId);
                }
                return _drawing;
            }
        }

        internal void LoadDocument()
        {
            //TODO:
            //check if file in .bin or other format

            var oleRel  = _worksheet.Part.GetRelationship(relId);
            var oleObj = UriHelper.ResolvePartUri(_worksheet.WorksheetUri, oleRel.TargetUri);
            var olePart = _worksheet._package.ZipPackage.GetPart(oleObj);
            var oleStream = (MemoryStream)olePart.GetStream(FileMode.Open, FileAccess.Read);
            _document = new CompoundDocument(oleStream);
        }

        internal void LoadExternalLink()
        {
            //_worksheet.Workbook.ExternalLinks
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
            return new OleObjectDrawing(oleObject, drawings, drawNode.ParentNode, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent);
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
