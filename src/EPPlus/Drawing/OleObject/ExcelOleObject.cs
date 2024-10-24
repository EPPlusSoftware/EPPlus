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
    /// Class for handling OLE Objects.
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
        internal static int ExternalLinkId = 1;

        /// <summary>
        /// True: File is displayed as Icon.
        /// </summary>
        public readonly bool DisplayAsIcon;

        /// <summary>
        /// True: File is linked. False: File is embedded.
        /// </summary>
        public readonly bool IsExternalLink;

        /// <summary>
        /// Returns the drawing type of this object.
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
        internal ExcelOleObject(ExcelDrawings drawings, XmlNode node, OleObjectInternal oleObject, ExcelGroupShape parent = null)
            : base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
        {
            _oleObject = oleObject;
            _worksheet = drawings.Worksheet;
            IsExternalLink = string.IsNullOrEmpty(_oleObject.Link);
            //Read vml
            _vml = drawings.Worksheet.VmlDrawings[LegacySpId];
            _vmlProp = XmlHelperFactory.Create(_vml.NameSpaceManager, _vml.GetNode("x:ClientData"));
            //read emf file uri
            var v = _oleObject.TopNode.ChildNodes[0].Attributes["r:id"].Value;
            var rel = _worksheet.Part.GetRelationship(v);
            var uri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            var emfStream = (MemoryStream)_worksheet._package.ZipPackage.GetPart(uri).GetStream();
            byte[] image = emfStream.ToArray();
            _mediaImage = _worksheet._package.PictureStore.AddImage(image, null, ePictureType.Emf);
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

        internal ExcelOleObject(ExcelDrawings drawings, XmlNode node, string name, string olePath, ExcelOleObjectParameters parameters, string iconPath = null, ExcelGroupShape parent = null)
            : base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
        {
            byte[] oleData = File.ReadAllBytes(olePath);
            parameters.Extension = Path.GetExtension(olePath);
            byte[] iconData;
            if(string.IsNullOrEmpty(iconPath))
            {
                iconData = null;
            }
            else
            {
                iconData = File.ReadAllBytes(iconPath);
                using MemoryStream ms = new MemoryStream(iconData);
                using BinaryReader br = new BinaryReader(ms);
                string sign;
                if(!ImageReader.IsBmp(br, out sign))
                {
                    throw new Exception("Invalid file format for Icon. Supported formats are .BMP");
                }
            }
            IsExternalLink = parameters.LinkToFile;
            DisplayAsIcon = parameters.DisplayAsIcon;
            CreateOleObject(drawings, node, Name, oleData, parameters, iconData, parent);
        }

        internal ExcelOleObject(ExcelDrawings drawings, XmlNode node, string name, FileInfo oleInfo, ExcelOleObjectParameters parameters, FileInfo iconInfo = null, ExcelGroupShape parent = null)
            : base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
        {
            byte[] oleData = null;
            parameters.Extension = string.IsNullOrEmpty(parameters.Extension) ? oleInfo.Extension : parameters.Extension;
            parameters.OlePath = oleInfo.FullName;
            if (parameters.LinkToFile == false)
            {
                using FileStream oleFs = oleInfo.OpenRead();
                {
                    oleData = new byte[oleFs.Length];
                    oleFs.Read(oleData, 0, oleData.Length);
                }
            }
            byte[] iconData;
            if (iconInfo == null)
            {
                iconData = null;
            }
            else
            {
                using FileStream icoFs = oleInfo.OpenRead();
                {
                    iconData = new byte[icoFs.Length];
                    icoFs.Read(iconData, 0, iconData.Length);
                }
            }
            IsExternalLink = parameters.LinkToFile;
            DisplayAsIcon = parameters.DisplayAsIcon;
            CreateOleObject(drawings, node, Name, oleData, parameters, iconData, parent);
        }

        internal ExcelOleObject(ExcelDrawings drawings, XmlNode node, string name, Stream oleStream, ExcelOleObjectParameters parameters, Stream iconStream = null, ExcelGroupShape parent = null)
            : base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
        {
            byte[] oleData = new byte[oleStream.Length];
            oleStream.Seek(0, SeekOrigin.Begin);
            oleStream.Read(oleData, 0, (int)oleStream.Length);
            byte[] iconData = null;
            if (iconStream != null)
            {
                iconData = new byte[iconStream.Length];
                iconStream.Seek(0, SeekOrigin.Begin);
                iconStream.Read(iconData, 0, (int)iconStream.Length);
            }
            IsExternalLink = parameters.LinkToFile;
            DisplayAsIcon = parameters.DisplayAsIcon;
            CreateOleObject(drawings, node, Name, oleData, parameters, iconData, parent);
        }

        internal void CreateOleObject(ExcelDrawings drawings, XmlNode node, string name, byte[] oleData, ExcelOleObjectParameters parameters, byte[] iconData = null, ExcelGroupShape parent = null)
        {
            _worksheet = drawings.Worksheet;
            string relId = "";
            string oleObjectNode = "";
            if (parameters.ProgId == null)
            {
                parameters.ProgId = GetProgId(parameters.Extension);
            }
            if (IsExternalLink)
            {
                if (parameters.OlePath == null)
                {
                    throw new Exception("Linked files requires a string file path argument.");
                }
                CreateLinkToObject(parameters.OlePath, parameters.ProgId);
                if (DisplayAsIcon)
                {
                    oleObjectNode = string.Format("<oleObject dvAspect=\"DVASPECT_ICON\" oleUpdate=\"OLEUPDATE_ONCALL\" progId=\"{0}\" link=\"[{1}]!''''\" shapeId=\"{2}\">", parameters.ProgId, ExternalLinkId, _id);
                }
                else
                {
                    oleObjectNode = string.Format("<oleObject oleUpdate=\"OLEUPDATE_ALWAYS\" progId=\"{0}\" link=\"[{1}]!''''\" shapeId=\"{2}\">", parameters.ProgId, ExternalLinkId, _id);
                }
            }
            else
            {
                relId = CreateEmbeddedObject(name, parameters, oleData);
                if (DisplayAsIcon)
                {
                    oleObjectNode = string.Format("<oleObject dvAspect=\"DVASPECT_ICON\" progId=\"{0}\" shapeId=\"{1}\" r:id=\"{2}\">", _oleDataStructures.CompObj.Reserved1.String, _id, relId);
                }
                else
                {
                    oleObjectNode = string.Format("<oleObject progId=\"{0}\" shapeId=\"{1}\" r:id=\"{2}\">", _oleDataStructures.CompObj.Reserved1.String, _id, relId);
                }

            }
            //Create Media
            byte[] image = OleObjectIcon.DefaultIcon;
            EmfImage emf = new EmfImage();
            emf.Read(image);
            if (iconData != null)
            {
                emf.ChangeImage(iconData);
            }
            else
            {
                var ext = parameters.Extension;
                if (ext.Contains("docx"))
                    emf.ChangeImage(OleObjectIcon.Docx_Icon_Bitmap);
                if (ext.Contains("pptx"))
                    emf.ChangeImage(OleObjectIcon.Pptx_Icon_Bitmap);
                if (ext.Contains("xlsx"))
                    emf.ChangeImage(OleObjectIcon.Xlsx_Icon_Bitmap);
                if (ext.Contains("pdf"))
                    emf.ChangeImage(OleObjectIcon.PDF_Icon_Bitmap);
            }
            if(parameters.OlePath == null)
            {
                emf.SetNewTextInDefaultEMFImage(name + parameters.Extension);
            }
            else
            {
                emf.SetNewTextInDefaultEMFImage(Path.GetFileName(parameters.OlePath));
            }
            image = emf.GetBytes();
            //Add image to Picture Store
            _mediaImage = _worksheet._package.PictureStore.AddImage(image, null, ePictureType.Emf);
            var imgRelId = _mediaImage.Part.CreateRelationship(_mediaImage.Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
            //Create drawings xml
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
            var wsNode = _worksheet.CreateOleContainerNode();
            StringBuilder sb = new StringBuilder();
            sb.Append("<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">");
            sb.Append("<mc:Choice Requires=\"x14\">");
            //Create object node
            sb.Append(oleObjectNode);
            if(IsExternalLink)
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
                        int linkId = int.Parse(splitFilename[0]);
                        ExternalLinkId = linkId > ExternalLinkId ? linkId : ExternalLinkId;
                        break;
                    }
                }
            }
        }

        private string CreateEmbeddedObject(string name, ExcelOleObjectParameters parameters, byte[] oleData)
        {
            string relId = "";
            _oleDataStructures = new OleObjectDataStructures();
            _document = new CompoundDocument();
            Guid ClsId = OleObjectGUIDCollection.keyValuePairs["Package"];
            if (parameters.Extension == ".pdf")
            {
                Ole.CreateOleObject(_oleDataStructures, IsExternalLink);
                Ole.CreateOleDataStream(_oleDataStructures, _document, IsExternalLink);
                CompObj.CreateCompObjObject(_oleDataStructures, "Acrobat Document", "Acrobat.Document.DC");
                CompObj.CreateCompObjDataStream(_oleDataStructures, _document);
                OleDataFile.CreateDataFileObject(_oleDataStructures, oleData);
                OleDataFile.CreateDataFileDataStream(_document, OleDataFile.CONTENTS_STREAM_NAME, oleData);
                ClsId = OleObjectGUIDCollection.keyValuePairs["PDF"];
            }
            else if (parameters.Extension == ".odp" || parameters.Extension == ".odt" || parameters.Extension == ".ods")
            {
                string UserType = "", Reserved = "", key = "";
                if (parameters.Extension == ".odp")
                {
                    UserType = "OpenDocument Presentation";
                    Reserved = "PowerPoint.OpenDocumentPresentation.12";
                    key = "ODP";
                }
                else if (parameters.Extension == ".odt")
                {
                    UserType = "OpenDocument Text";
                    Reserved = "Word.OpenDocumentText.12";
                    key = "ODT";
                }
                else if (parameters.Extension == ".ods")
                {
                    UserType = "OpenDocument Spreadsheet";
                    Reserved = "Excel.OpenDocumentSpreadsheet.12";
                    key = "ODS";
                }
                Ole.CreateOleObject(_oleDataStructures, IsExternalLink);
                Ole.CreateOleDataStream(_oleDataStructures, _document, IsExternalLink);
                OleDataFile.CreateDataFileObject(_oleDataStructures, oleData);
                OleDataFile.CreateDataFileDataStream(_document, OleDataFile.EMBEDDEDODF_STREAM_NAME, oleData);
                CompObj.CreateCompObjObject(_oleDataStructures, UserType, Reserved);
                CompObj.CreateCompObjDataStream(_oleDataStructures, _document);
                ClsId = OleObjectGUIDCollection.keyValuePairs[key];
            }
            else if (parameters.Extension == ".docx" || parameters.Extension == ".pptx" || parameters.Extension == ".xlsx")
            {
                //Embedd as is
                string oleName = "";
                string contentType = "";
                string ext = "";
                if (parameters.Extension == ".docx")
                {
                    oleName = "Microsoft_Word_Document";
                    contentType = ContentTypes.contentTypeOleDocx;
                    ext = "docx";
                }
                else if (parameters.Extension == ".xlsx")
                {
                    oleName = "Microsoft_Excel_Worksheet";
                    contentType = ContentTypes.contentTypeOleXlsx;
                    ext = "xlsx";
                }
                else if (parameters.Extension == ".pptx")
                {
                    oleName = "Microsoft_PowerPoint_Presentation";
                    contentType = ContentTypes.contentTypeOlePptx;
                    ext = "pptx";
                }
                int newID = 1;
                var Uri = GetNewUri(_worksheet._package.ZipPackage, "/xl/embeddings/" + oleName + "{0}" + parameters.Extension, ref newID);
                _oleObjectPart = _worksheet._package.ZipPackage.CreatePart(Uri, contentType, CompressionLevel.None, ext);
                var rel = _worksheet.Part.CreateRelationship(Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/package");
                relId = rel.Id;
                MemoryStream ms = (MemoryStream)_oleObjectPart.GetStream(FileMode.Create, FileAccess.Write);
                ms.Write(oleData, 0, oleData.Length);
                return relId;
            }
            else
            {
                string oleName = "";
                if(parameters.OlePath == null)
                {
                    oleName = name + parameters.Extension;
                }
                else
                {
                    oleName = parameters.OlePath;
                }
                CompObj.CreateCompObjObject(_oleDataStructures, "OLE Package", "Package");
                CompObj.CreateCompObjDataStream(_oleDataStructures, _document);
                Ole10Native.CreateOle10NativeObject(oleData, oleName, _oleDataStructures);
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

        private void CreateLinkToObject(string filePath, string progId)
        {
            var wb = _worksheet.Workbook;
            //create externalLink xml part
            Uri uri = GetNewUri(wb._package.ZipPackage, "/xl/externalLinks/externalLink{0}.xml", ref ExternalLinkId);
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
            xml.AppendFormat(" r:id=\"{0}\" progId=\"{1}\">", fileRel.Id, progId);
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
        }

        private string GetProgId(string extension)
        {
            switch (extension)
            {
                case ".pdf":
                    return "Acrobat.Document.DC";
                case ".docx":
                    return "Word.Document.12";
                case ".xlsx":
                    return "Excel.Sheet.12";
                case ".pptx":
                    return "PowerPoint.Show.12";
                case ".ods":
                    return "Excel.OpenDocumentSpreadsheet.12";
                case ".odt":
                    return "Word.OpenDocumentText.12";
                case ".odp":
                    return"PowerPoint.OpenDocumentPresentation.12";
                default:
                    return "Package";
            }
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
        /// Internal Method for debugging purposes. Exports the contents of the data streams inside a compound document containing an OLE Object.
        /// </summary>
        /// <param name="exportPath"></param>
        internal void ExportOleObjectData(string exportPath)
        {
            if (IsExternalLink || _document == null)
                return;
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
            using var p = new ExcelPackage(exportPath);
            OleObjectDataStreamsExport.ExportOleNative(_worksheet._package.File.Name, _oleObjectPart.Entry.FileName, p, _oleDataStructures);
            OleObjectDataStreamsExport.ExportOle(_worksheet._package.File.Name, _oleObjectPart.Entry.FileName, p, _oleDataStructures, IsExternalLink);
            OleObjectDataStreamsExport.ExportCompObj(_worksheet._package.File.Name, _oleObjectPart.Entry.FileName, p, _oleDataStructures);
            p.Save();
        }
        #endregion
    }
}
