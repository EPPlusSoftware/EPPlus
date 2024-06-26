using System;
using System.Xml;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Utils.CompundDocument;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.Utils;
using System.IO;

namespace OfficeOpenXml.Drawing.OleObject
{
    internal class ExcelOleObject : ExcelDrawing
    {
        internal ExcelVmlDrawingBase _vml;
        internal XmlHelper _vmlProp;
        internal OleObjectInternal _oleObject;
        internal CompoundDocument _document;
        internal ExcelExternalOleLink _externalLink;
        internal ExcelWorksheet _worksheet;
        public bool isExternalLink = false;
        internal ExcelOleObject(ExcelDrawings drawings, XmlNode node, OleObjectInternal oleObject, ExcelGroupShape parent = null) : base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
        {
            _oleObject = oleObject;
            _worksheet = drawings.Worksheet;

            _vml = drawings.Worksheet.VmlDrawings[LegacySpId];
            _vmlProp = XmlHelperFactory.Create(_vml.NameSpaceManager, _vml.GetNode("x:ClientData"));

            if (string.IsNullOrEmpty(_oleObject.Link))
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

        internal void LoadDocument()
        {
            //TODO:
            //check if file in .bin or other format

            var oleRel = _worksheet.Part.GetRelationship(_oleObject.RelationshipId);
            var oleObj = UriHelper.ResolvePartUri(oleRel.SourceUri, oleRel.TargetUri);
            var olePart = _worksheet._package.ZipPackage.GetPart(oleObj);
            var oleStream = (MemoryStream)olePart.GetStream(FileMode.Open, FileAccess.Read);
            _document = new CompoundDocument(oleStream);
        }

        internal void LoadExternalLink()
        {
            var els = _worksheet.Workbook.ExternalLinks;
            foreach (var el in els)
            {
                if (el.ExternalLinkType == eExternalLinkType.OleLink)
                {
                    var filename = el.Part.Entry.ToString();
                    var splitFilename = filename.Split("ZipEntry::xl/externalLinks/externalLink.xml".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                    var splitLink = _oleObject.Link.Split("[]".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                    if (splitLink[0].Contains(splitFilename[0]))
                    {
                        _externalLink = el as ExcelExternalOleLink;
                        break;
                    }
                }
            }
        }

        public override eDrawingType DrawingType
        {
            get
            {
                return eDrawingType.OleObject;
            }
        }

        internal string LegacySpId
        {
            get
            {
                return GetXmlNodeString($"{GetlegacySpIdPath()}/a:extLst/a:ext[@uri='{ExtLstUris.LegacyObjectWrapperUri}']/a14:compatExt/@spid");
            }
            set
            {
                var node = GetNode(GetlegacySpIdPath());
                var extHelper = XmlHelperFactory.Create(NameSpaceManager, node);
                var extNode = extHelper.GetOrCreateExtLstSubNode(ExtLstUris.LegacyObjectWrapperUri, "a14");
                if (extNode.InnerXml == "")
                {
                    extNode.InnerXml = $"<a14:compatExt/>";
                }
                ((XmlElement)extNode.FirstChild).SetAttribute("spid", value);
            }
        }
        internal string GetlegacySpIdPath()
        {
            return $"{(_topPath == "" ? "" : _topPath + "/")}xdr:nvSpPr/xdr:cNvPr";
        }
    }
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