using System;
using System.Xml;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Utils.CompundDocument;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System.IO;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System.Xml.Linq;
using OfficeOpenXml.Utils.Extensions;

namespace OfficeOpenXml.Drawing.OleObject
{
    public class ExcelOleObject : ExcelDrawing
    {
        internal ExcelVmlDrawingBase _vml;
        internal XmlHelper _vmlProp;
        internal OleObjectInternal _oleObject;
        internal CompoundDocument _document;
        internal ExcelExternalOleLink _externalLink;
        internal ExcelWorksheet _worksheet;
        public bool isExternalLink = false;
        internal ExcelOleObject(ExcelDrawings drawings, XmlNode node, OleObjectInternal oleObject, ExcelGroupShape parent = null)
            : base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
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

        internal ExcelOleObject(ExcelDrawings drawings, XmlNode node, string filepath, bool link, string mediaFilePath = "", ExcelGroupShape parent = null)
            : base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
        {
            _worksheet = drawings.Worksheet;


            //Create this first and check if successful creation before creating xml for other parts
            //Create ExternalLink
            //       OR
            //Create Embedded Document
            //.bin
            //.doc
            //osv
            string relId = "";
            if (link)
            {
                isExternalLink = true;
                //create ExternalLink
            }
            else
            {
                isExternalLink= false;
                //create embedded object

                //Create embeddingsfolder
                int newID = 1;
                var Uri = GetNewUri(_worksheet._package.ZipPackage, "/xl/embeddings/oleObject{0}.xml", ref newID);
                var part = _worksheet._package.ZipPackage.CreatePart(Uri, ContentTypes.contentTypeControlProperties);
                var rel = _worksheet.Part.CreateRelationship(Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/embeddings");

                using (MemoryStream ms = new MemoryStream())
                {
                    byte[] data = File.ReadAllBytes(filepath);
                    ms.Write(data, 0, data.Length);
                    _document = new CompoundDocument();
                    _document.Save(ms);
                }
            }

            //Create Media
            //User supplied picture or our own placeholder
            //Construct icon with rectable with txbody set to filename and an autorectangle. Somehow you can't see the txbody or autorectangle when icon is complete. only when you ungroup.

            //Create drawings xml
            XmlElement spElement = CreateShapeNode();
            spElement.InnerXml = CreateOleObjectDrawingNode();
            CreateClientData();

            //Create vml
            _vml = drawings.Worksheet.VmlDrawings.AddPicture(this, _drawings.GetUniqueDrawingName("Object 1"));
            _vmlProp = XmlHelperFactory.Create(_vml.NameSpaceManager, _vml.GetNode("x:ClientData"));

            //Create worksheet xml
            //Create collection container node
            var wsNode = _worksheet.CreateOleContainerNode();
            StringBuilder sb = new StringBuilder();
            sb.Append("<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">");
            sb.Append("<mc:Choice Requires=\"x14\">");

            //Create object node
            sb.AppendFormat("<oleObject progId=\"Acrobat.Document.DC\" shapeId=\"{0}\" r:id=\"{1}\">", _id, "obj"); //SET relId TO EMBEDDED/LINKED OBJECT
            sb.Append("<objectPr defaultSize=\"0\" autoPict=\"0\">"); //SET relId TO MEDIA HERE
            sb.Append("<anchor moveWithCells=\"1\">");
            sb.Append("<from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></from>");       //SET VALUE BASED ON MEDIA
            sb.Append("<to><xdr:col>1</xdr:col><xdr:colOff>304800</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>114300</xdr:rowOff></to>"); //SET VALUE BASED ON MEDIA
            sb.Append("</anchor></objectPr></oleObject>");

            sb.Append("</mc:Choice>");
            //fallback
            sb.AppendFormat("<mc:Fallback><oleObject progId=\"Acrobat.Document.DC\" shapeId=\"{0}\" r:id=\"{1}\" />", _id, "obj"); //SET relId TO EMBEDDED/LINKED OBJECT

            sb.Append("</mc:Fallback></mc:AlternateContent>");

            wsNode.InnerXml = sb.ToString();
            var oleObjectNode = wsNode.GetChildAtPosition(0).GetChildAtPosition(0);
            _oleObject = new OleObjectInternal(_worksheet.NameSpaceManager, oleObjectNode);
        }

        private string CreateOleObjectDrawingNode()
        {
            StringBuilder xml = new StringBuilder();
            xml.Append($"<xdr:nvSpPr><xdr:cNvPr hidden=\"1\" name=\"\" id=\"{_id}\"><a:extLst><a:ext uri=\"{{63B3BB69-23CF-44E3-9099-C40C66FF867C}}\"><a14:compatExt spid=\"_x0000_s{_id}\"/></a:ext><a:ext uri=\"{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}\"><a16:creationId id=\"{{00000000-0008-0000-0000-000001040000}}\" xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>");
            xml.Append($"<xdr:spPr bwMode=\"auto\"><a:xfrm><a:off y=\"0\" x=\"0\"/><a:ext cy=\"0\" cx=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
            xml.Append($"<a:solidFill><a:srgbClr val=\"FFFFFF\" mc:Ignorable=\"a14\" a14:legacySpreadsheetColorIndex=\"65\"/></a:solidFill><a:ln w=\"9525\"><a:solidFill><a:srgbClr val=\"000000\" mc:Ignorable=\"a14\" a14:legacySpreadsheetColorIndex=\"64\"/></a:solidFill><a:miter lim=\"800000\"/><a:headEnd/><a:tailEnd/></a:ln></xdr:spPr>");
            return xml.ToString();
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

/*
 * OLE objekt 
 * Worksheet:
 *  relId -> drawing
 *  relId -> legacyDrawing
 *  oleobject/relId -> embedding
 *  oleobject/link  -> externalLink
 *  oleobject/objectPr/relId -> media
 *
 * Drawing:
 *  sp/cNvPr/id -> vml
 *
 * VML:
 *  Samma id från Drawing
 *  relId -> media
 *
 * Embeddings:
 *  bin fil -> compound document
 *      har 3 filer, CONTENT (själva dokumentet, video, exe eller whatever), ole, CompObj
 *  doc filer och liknande ligger löst
 *      Microsoft_Word_Document, Microsoft_Word_Document1
 *
 * ExternalLinks:
 *  relId -> File Path
 *  Verkar som att siffran i filnamnet är länkad med siffran i worksheet/oleobject/link
 *  Har relation från workbook.xml
 *
 * Media:
 *  bild på .emf format
 *
 * PrinterSettings:
 *  bin file
 *  not supported
 */