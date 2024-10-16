using System;
using System.Xml;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Utils.CompundDocument;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Utils;
using System.IO;
using System.Text;
using OfficeOpenXml.Utils.Extensions;
using OfficeOpenXml.Drawing.OleObject.Structures;


namespace OfficeOpenXml.Drawing.OleObject
{
    /// <summary>
    /// Class for reading and writing OLE Objects.
    /// </summary>
    public class ExcelOleObject : ExcelDrawing
    {
        internal ExcelVmlDrawingBase _vml;
        internal XmlHelper _vmlProp;
        internal OleObjectInternal _oleObject;
        internal CompoundDocument _document;
        internal OleObjectDataStructures _oleDataStructures;
        internal ExcelExternalOleLink _externalLink;
        internal ExcelWorksheet _worksheet;
        internal ZipPackagePart _oleObjectPart;
        internal XmlDocument _linkedOleObjectXml;
        internal string _linkedObjectFilepath;
        internal ImageInfo _mediaImage;

        /// <summary>
        /// True: File is displayed as Icon.
        /// </summary>
        public readonly bool DisplayAsIcon;

        /// <summary>
        /// True: File is linked. False: File is embedded.
        /// </summary>
        public readonly bool IsExternalLink;

        /// <summary>
        /// Return the drawing type of this object
        /// </summary>
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

        /// <summary>
        /// Constructor for loading exsisting Ole Object.
        /// </summary>
        /// <param name="drawings"></param>
        /// <param name="node"></param>
        /// <param name="oleObject"></param>
        /// <param name="parent"></param>
        internal ExcelOleObject(ExcelDrawings drawings, XmlNode node, OleObjectInternal oleObject, ExcelGroupShape parent = null)
            : base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
        {
            _oleObject = oleObject;
            _worksheet = drawings.Worksheet;
            IsExternalLink = string.IsNullOrEmpty(_oleObject.Link);

            _vml = drawings.Worksheet.VmlDrawings[LegacySpId];
            _vmlProp = XmlHelperFactory.Create(_vml.NameSpaceManager, _vml.GetNode("x:ClientData"));

            //read emf file uri
            var v = _oleObject.TopNode.ChildNodes[0].Attributes["r:id"].Value;
            var rel = _worksheet.Part.GetRelationship(v);
            _mediaImage = new ImageInfo();
            _mediaImage.Uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);

            if (IsExternalLink)
            {
                IsExternalLink = false;
                LoadEmbeddedObject();
            }
            else
            {
                IsExternalLink = true;
                LoadLinkedObject();
            }
        }

        /// <summary>
        /// Constructor for creating new Ole Object.
        /// </summary>
        /// <param name="drawings"></param>
        /// <param name="node"></param>
        /// <param name="filePath"></param>
        /// <param name="linkToFile"></param>
        /// <param name="displayAsIcon"></param>
        /// <param name="iconFilePath"></param>
        /// <param name="parent"></param>
        internal ExcelOleObject(ExcelDrawings drawings, XmlNode node, string filePath, bool linkToFile, bool displayAsIcon = false, string iconFilePath = "", ExcelGroupShape parent = null)
            : base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
        {
            _worksheet = drawings.Worksheet;
            string relId = "";
            string oleObjectNode = "";
            IsExternalLink = linkToFile;
            DisplayAsIcon = displayAsIcon;
            if (linkToFile)
            {
                var linkId = CreateLinkToObject(filePath);
                if (displayAsIcon)
                {
                    oleObjectNode = string.Format("<oleObject dvAspect=\"DVASPECT_ICON\" oleUpdate=\"OLEUPDATE_ONCALL\" progId=\"{0}\" link=\"[{1}]!''''\" shapeId=\"{2}\">", "Package", linkId, _id);
                }
                else
                {
                    oleObjectNode = string.Format("<oleObject oleUpdate=\"OLEUPDATE_ALWAYS\" progId=\"{0}\" link=\"[{1}]!''''\" shapeId=\"{2}\">", "Package", linkId, _id);
                }
            }
            else
            {
                relId = CreateEmbeddedObject(filePath);
                if (displayAsIcon)
                {
                    oleObjectNode = string.Format("<oleObject dvAspect=\"DVASPECT_ICON\" progId=\"{0}\" shapeId=\"{1}\" r:id=\"{2}\">", _oleDataStructures.CompObj.Reserved1.String, _id, relId);
                }
                else
                {
                    oleObjectNode = string.Format("<oleObject progId=\"{0}\" shapeId=\"{1}\" r:id=\"{2}\">", _oleDataStructures.CompObj.Reserved1.String, _id, relId);
                }

            }
            //Create Media
            //int newID = 1;
            //_mediaUri = GetNewUri(_worksheet._package.ZipPackage, "/xl/media/image{0}.emf", ref newID);
            //var part = _worksheet._package.ZipPackage.CreatePart(_mediaUri, "image/x-emf", CompressionLevel.None, "emf");
            //var rel = _worksheet.Part.CreateRelationship(_mediaUri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
            byte[] image = OleObjectIcon.DefaultIcon;
            EmfImage emf = new EmfImage();
            emf.Read(image);
            if (!string.IsNullOrEmpty(iconFilePath))
            {
                byte[] newImage = File.ReadAllBytes(iconFilePath);
                emf.ChangeImage(newImage);
            }
            else
            {
                var ext = Path.GetExtension(filePath).ToLower();
                if (ext.Contains("docx"))
                    emf.ChangeImage(OleObjectIcon.Docx_Icon_Bitmap);
                if (ext.Contains("pptx"))
                    emf.ChangeImage(OleObjectIcon.Pptx_Icon_Bitmap);
                if (ext.Contains("xlsx"))
                    emf.ChangeImage(OleObjectIcon.Xlsx_Icon_Bitmap);
                if (ext.Contains("pdf"))
                    emf.ChangeImage(OleObjectIcon.PDF_Icon_Bitmap);
            }
            string filename = Path.GetFileName(filePath);
            emf.SetNewTextInDefaultEMFImage(filename);
            image = emf.GetBytes();

            //use PictureStore here?
            _mediaImage = _worksheet._package.PictureStore.AddImage(image, null, ePictureType.Emf);
            var imgRelId = _mediaImage.Part.CreateRelationship(_mediaImage.Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");

            //MemoryStream ms = (MemoryStream)part.GetStream(FileMode.Create, FileAccess.Write);
            //ms.Write(image, 0, image.Length);

            //Create drawings xml
            string name = _drawings.GetUniqueDrawingName("Object 1");
            XmlElement spElement = CreateShapeNode();
            spElement.InnerXml = CreateOleObjectDrawingNode(name);
            CreateClientData();
            From.Column = 0;  From.ColumnOff = 0;
            From.Row = 0;     From.RowOff = 0;
            To.Column = 1;    To.ColumnOff = 304800;
            To.Row = 3;       To.RowOff = 114300;

            //Create vml
            _vml = drawings.Worksheet.VmlDrawings.AddOlePicture(this.Id.ToString(), imgRelId.TargetUri);
            _vmlProp = XmlHelperFactory.Create(_vml.NameSpaceManager, _vml.GetNode("x:ClientData"));

            //Create worksheet xml
            //Create collection container node
            var wsNode = _worksheet.CreateOleContainerNode();
            StringBuilder sb = new StringBuilder();
            sb.Append("<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">");
            sb.Append("<mc:Choice Requires=\"x14\">");
            //Create object node
            sb.Append(oleObjectNode);
            if(linkToFile)
                sb.AppendFormat("<objectPr defaultSize=\"0\" r:id=\"{0}\" dde=\"1\">", imgRelId.Id);
            else
                sb.AppendFormat("<objectPr defaultSize=\"0\" r:id=\"{0}\">", imgRelId.Id);
            sb.Append("<anchor moveWithCells=\"1\">");
            sb.AppendFormat("<from><xdr:col>{0}</xdr:col><xdr:colOff>{1}</xdr:colOff><xdr:row>{2}</xdr:row><xdr:rowOff>{3}</xdr:rowOff></from>", From.Column, From.ColumnOff, From.Row, From.RowOff);
            sb.AppendFormat("<to><xdr:col>{0}</xdr:col><xdr:colOff>{1}</xdr:colOff><xdr:row>{2}</xdr:row><xdr:rowOff>{3}</xdr:rowOff></to>", To.Column, To.ColumnOff, To.Row, To.RowOff);
            sb.Append("</anchor></objectPr></oleObject>");
            sb.Append("</mc:Choice>");
            //fallback
            sb.AppendFormat("<mc:Fallback>");
            sb.Append(oleObjectNode + "</oleObject>");
            sb.Append("</mc:Fallback></mc:AlternateContent>");
            wsNode.InnerXml = sb.ToString();
            var oleObjectXmlNode = wsNode.GetChildAtPosition(0).GetChildAtPosition(0);
            _oleObject = new OleObjectInternal(_worksheet.NameSpaceManager, oleObjectXmlNode);
        }

        private string CreateOleObjectDrawingNode(string name)
        {
            StringBuilder xml = new StringBuilder();
            xml.Append($"<xdr:nvSpPr>" +
                       $"<xdr:cNvPr hidden=\"1\" name=\"{name}\" id=\"{_id}\">" +
                       $"<a:extLst>" +
                       $"<a:ext uri=\"{{63B3BB69-23CF-44E3-9099-C40C66FF867C}}\">" +
                       $"<a14:compatExt spid=\"_x0000_s{_id}\"/>" +
                       $"</a:ext>" +
                       $"<a:ext uri=\"{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}\">" +
                       $"<a16:creationId id=\"{{C4F0F4B0-B1B7-3F07-7766-FB369B01C1A5}}\" xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\"/>" +
                       $"</a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>");
            xml.Append($"<xdr:spPr bwMode=\"auto\">" +
                       $"<a:xfrm>" +
                       $"<a:off y=\"0\" x=\"0\"/>" +
                       $"<a:ext cy=\"0\" cx=\"0\"/>" +
                       $"</a:xfrm>" +
                       $"<a:prstGeom prst=\"rect\">" +
                       $"<a:avLst/></a:prstGeom>");
            xml.Append($"<a:solidFill>" +
                       $"<a:srgbClr val=\"FFFFFF\" mc:Ignorable=\"a14\" a14:legacySpreadsheetColorIndex=\"65\"/>" +
                       $"</a:solidFill><a:ln w=\"9525\">" +
                       $"<a:solidFill>" +
                       $"<a:srgbClr val=\"000000\" mc:Ignorable=\"a14\" a14:legacySpreadsheetColorIndex=\"64\"/>" +
                       $"</a:solidFill>" +
                       $"<a:prstDash val=\"solid\"/>" +
                       $"<a:miter lim=\"800000\"/>" +
                       $"<a:headEnd/>" +
                       $"<a:tailEnd type=\"none\" w=\"med\" len=\"med\"/>" +
                       $"</a:ln>");
            xml.Append($"<a:effectLst/><a:extLst>" +
                       $"<a:ext uri=\"{{AF507438-7753-43E0-B8FC-AC1667EBCBE1}}\">" +
                       $"<a14:hiddenEffects>" +
                       $"<a:effectLst>" +
                       $"<a:outerShdw dist=\"35921\" dir=\"2700000\" algn=\"ctr\" rotWithShape=\"0\">" +
                       $"<a:srgbClr val=\"808080\" />" +
                       $"</a:outerShdw></a:effectLst></a14:hiddenEffects></a:ext></a:extLst></xdr:spPr>");
            return xml.ToString();
        }

        private void LoadEmbeddedObject()
        {
            var oleRel = _worksheet.Part.GetRelationship(_oleObject.RelationshipId);
            if (oleRel != null && oleRel.TargetUri.ToString().Contains(".bin"))
            {
                var oleObj = UriHelper.ResolvePartUri(oleRel.SourceUri, oleRel.TargetUri);
                _oleObjectPart = _worksheet._package.ZipPackage.GetPart(oleObj);
                var oleStream = (MemoryStream)_oleObjectPart.GetStream(FileMode.Open, FileAccess.Read);
                _document = new CompoundDocument(oleStream);
            }
            else if(oleRel != null && ( oleRel.TargetUri.ToString().Contains(".docx") || oleRel.TargetUri.ToString().Contains(".pptx") || oleRel.TargetUri.ToString().Contains(".xlsx")))
            {
                var oleObj = UriHelper.ResolvePartUri(oleRel.SourceUri, oleRel.TargetUri);
                _oleObjectPart = _worksheet._package.ZipPackage.GetPart(oleObj);
            }
        }

        private void LoadLinkedObject()
        {
            var els = _worksheet.Workbook.ExternalLinks;
            foreach (var el in els)
            {
                if (el.ExternalLinkType == eExternalLinkType.OleLink)
                {
                    var filename = el.Part.Uri.ToString();
                    var splitFilename = filename.Split("/xl/externalLinks/externalLink.xml".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                    var splitLink = _oleObject.Link.Split("[]".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                    if (splitLink[0].Contains(splitFilename[0]))
                    {
                        _externalLink = el as ExcelExternalOleLink;
                        _linkedOleObjectXml = _externalLink.ExternalOleXml;
                        _linkedObjectFilepath = _externalLink.Relation.TargetUri.OriginalString;
                        break;
                    }
                }
            }
        }

        private string CreateEmbeddedObject(string filePath)
        {
            string relId = "";
            byte[] fileData = File.ReadAllBytes(filePath);
            string fileType = Path.GetExtension(filePath).ToLower();
            _oleDataStructures = new OleObjectDataStructures();
            _document = new CompoundDocument();
            Guid ClsId = OleObjectGUIDCollection.keyValuePairs["Package"];
            if (fileType == ".pdf")
            {
                Ole.CreateOleObject(_oleDataStructures, IsExternalLink);
                Ole.CreateOleDataStream(_oleDataStructures, _document, IsExternalLink);
                CompObj.CreateCompObjObject(_oleDataStructures, "Acrobat Document", "Acrobat.Document.DC");
                CompObj.CreateCompObjDataStream(_oleDataStructures, _document);
                OleDataFile.CreateDataFileObject(_oleDataStructures, fileData);
                OleDataFile.CreateDataFileDataStream(_document, OleDataFile.CONTENTS_STREAM_NAME, fileData);
                ClsId = OleObjectGUIDCollection.keyValuePairs["PDF"];
            }
            else if (fileType == ".odp" || fileType == ".odt" || fileType == ".ods")
            {
                string UserType = "", Reserved = "", key = "";
                if (fileType == ".odp")
                {
                    UserType = "OpenDocument Presentation";
                    Reserved = "PowerPoint.OpenDocumentPresentation.12";
                    key = "ODP";
                }
                else if (fileType == ".odt")
                {
                    UserType = "OpenDocument Text";
                    Reserved = "Word.OpenDocumentText.12";
                    key = "ODT";
                }
                else if (fileType == ".ods")
                {
                    UserType = "OpenDocument Spreadsheet";
                    Reserved = "Excel.OpenDocumentSpreadsheet.12";
                    key = "ODS";
                }
                Ole.CreateOleObject(_oleDataStructures, IsExternalLink);
                Ole.CreateOleDataStream(_oleDataStructures, _document, IsExternalLink);
                OleDataFile.CreateDataFileObject(_oleDataStructures, fileData);
                OleDataFile.CreateDataFileDataStream(_document, OleDataFile.EMBEDDEDODF_STREAM_NAME, fileData);
                CompObj.CreateCompObjObject(_oleDataStructures, UserType, Reserved);
                CompObj.CreateCompObjDataStream(_oleDataStructures, _document);
                ClsId = OleObjectGUIDCollection.keyValuePairs[key];
            }
            else if (fileType == ".docx" || fileType == ".pptx" || fileType == ".xlsx")
            {
                //Embedd as is
                string name = "";
                string contentType = "";
                string ext = "";
                if (fileType == ".docx")
                {
                    name = "Microsoft_Word_Document";
                    CompObj.CreateCompObjObject(_oleDataStructures, "Document", "Document");
                    
                    ext = "docx";
                }
                else if (fileType == ".xlsx")
                {
                    name = "Microsoft_Excel_Worksheet";
                    CompObj.CreateCompObjObject(_oleDataStructures, "Worksheet", "Worksheet");
                    contentType = ContentTypes.contentTypeOleXlsx;
                    ext = "xlsx";
                }
                else if (fileType == ".pptx")
                {
                    name = "Microsoft_PowerPoint_Presentation";
                    CompObj.CreateCompObjObject(_oleDataStructures, "Presentation", "Presentation");
                    contentType = ContentTypes.contentTypeOlePptx;
                    ext = "pptx";
                }
                int newID = 1;
                var Uri = GetNewUri(_worksheet._package.ZipPackage, "/xl/embeddings/" + name + "{0}" + fileType, ref newID);
                _oleObjectPart = _worksheet._package.ZipPackage.CreatePart(Uri, contentType, CompressionLevel.None, ext);
                var rel = _worksheet.Part.CreateRelationship(Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/package");
                relId = rel.Id;
                MemoryStream ms = (MemoryStream)_oleObjectPart.GetStream(FileMode.Create, FileAccess.Write);
                ms.Write(fileData, 0, fileData.Length);
                return relId;
            }
            else
            {
                CompObj.CreateCompObjObject(_oleDataStructures, "OLE Package", "Package");
                CompObj.CreateCompObjDataStream(_oleDataStructures, _document);
                Ole10Native.CreateOle10NativeObject(fileData, filePath, _oleDataStructures);
                Ole10Native.CreateOle10NativeDataStream(_oleDataStructures, _document);
                ClsId = OleObjectGUIDCollection.keyValuePairs["Package"];
            }
            if (_document.Storage.DataStreams != null)
            {
                int newID = 1;
                var Uri = GetNewUri(_worksheet._package.ZipPackage, "/xl/embeddings/oleObject{0}.bin", ref newID);
                _oleObjectPart = _worksheet._package.ZipPackage.CreatePart(Uri, ContentTypes.contentTypeOleObject);
                var rel = _worksheet.Part.CreateRelationship(Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/oleObject");
                MemoryStream ms = (MemoryStream)_oleObjectPart.GetStream(FileMode.Create, FileAccess.Write);
                _document.RootItem.ClsID = ClsId;
                _document.Save(ms);
                relId = rel.Id;
            }
            return relId;
        }

        private int CreateLinkToObject(string filePath)
        {
            var wb = _worksheet.Workbook;
            //create externalLink xml part
            int newID = 1;
            Uri uri = GetNewUri(wb._package.ZipPackage, "/xl/externalLinks/externalLink{0}.xml", ref newID);
            _oleObjectPart = wb._package.ZipPackage.CreatePart(uri, ContentTypes.contentTypeExternalLink);
            var rel = wb.Part.CreateRelationship(uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/externalLink");
            //Create relation to external file
            _linkedObjectFilepath = "file:///" + filePath;
            var fileRel = _oleObjectPart.CreateRelationship(_linkedObjectFilepath, TargetMode.External, ExcelPackage.schemaRelationships + "/oleObject");
            //Create externalLink xml
            var xml = new StringBuilder();
            xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            xml.Append("<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
            xml.Append(" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14 xxl21\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"");
            xml.Append(" xmlns:xxl21=\"http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021\">");
            xml.Append("<oleLink xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
            xml.AppendFormat(" r:id=\"{0}\" progId=\"{1}\">", fileRel.Id, "Package");
            if (DisplayAsIcon)
                xml.AppendFormat("<oleItems><oleItem name=\"{0}\" icon=\"{1}\" preferPic=\"{2}\"/>", "\'", "1", "1");
            else
                xml.AppendFormat("<oleItems><oleItem name=\"{0}\" advise=\"{1}\" preferPic=\"{2}\"/>", "\'", "1", "1");
            xml.Append("</oleItems></oleLink></externalLink>");
            _linkedOleObjectXml = new XmlDocument();
            _linkedOleObjectXml.LoadXml(xml.ToString());
            _linkedOleObjectXml.Save(_oleObjectPart.GetStream(FileMode.Create, FileAccess.Write));

            //create/write wb xml external link node
            var er = (XmlElement)wb.CreateNode("d:externalReferences/d:externalReference", false, true);
            er.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);

            //Add the externalLink to externalLink collection
            _externalLink = wb.ExternalLinks[wb.ExternalLinks.GetExternalLink(filePath, fileRel)] as ExcelExternalOleLink;
            return newID;
        }

        internal override void DeleteMe()
        {
            if (IsExternalLink)
            {
                //delete externalReferences
                _worksheet.Workbook.ExternalLinks.Remove(_externalLink);
            }
            else
            {
                //Delete embeddings
                _worksheet._package.ZipPackage.DeletePart(_oleObjectPart.Uri);
            }
            //Delete vml
            _vml.TopNode.ParentNode.RemoveChild(_vml.TopNode);
            //Delete worksheet & Internal Representation
            _oleObject.DeleteMe();
            //delete media
            _worksheet._package.ZipPackage.DeletePart(_mediaImage.Uri);
            //Delete drawing
            base.DeleteMe();
        }

        #region Export
        /// <summary>
        /// Internal Method for debugging purposes.
        /// </summary>
        /// <param name="ExportPath"></param>
        internal void ExportOleObjectData(string ExportPath)
        {
            _oleDataStructures = new OleObjectDataStructures();
            if (_document.Storage.DataStreams.ContainsKey(Ole10Native.OLE10NATIVE_STREAM_NAME))
            {
                _oleDataStructures.OleNative = new OleObjectDataStructures.OleNativeStream();
                Ole10Native.ReadOle10Native(_oleDataStructures, _document.Storage.DataStreams[Ole10Native.OLE10NATIVE_STREAM_NAME].Stream);
            }
            if (_document.Storage.DataStreams.ContainsKey(Ole.OLE_STREAM_NAME))
            {
                _oleDataStructures.Ole = new OleObjectDataStructures.OleObjectStream();
                Ole.ReadOleStream(_oleDataStructures, _document.Storage.DataStreams[Ole.OLE_STREAM_NAME].Stream);
            }
            if (_document.Storage.DataStreams.ContainsKey(CompObj.COMPOBJ_STREAM_NAME))
            {
                _oleDataStructures.CompObj = new OleObjectDataStructures.CompObjStream();
                CompObj.ReadCompObjStream(_oleDataStructures, _document.Storage.DataStreams[CompObj.COMPOBJ_STREAM_NAME].Stream);
            }
            using var p = new ExcelPackage(ExportPath);
            OleObjectDataStreamsExport.ExportOleNative(_worksheet._package.File.Name, _oleObjectPart.Entry.FileName, p, _oleDataStructures);
            OleObjectDataStreamsExport.ExportOle(_worksheet._package.File.Name, _oleObjectPart.Entry.FileName, p, _oleDataStructures, IsExternalLink);
            OleObjectDataStreamsExport.ExportCompObj(_worksheet._package.File.Name, _oleObjectPart.Entry.FileName, p, _oleDataStructures);
            p.Save();
        }
        #endregion
    }
}
