﻿using System;
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
using static OfficeOpenXml.Drawing.OleObject.OleObjectDataStructures;
using System.Collections.Generic;


namespace OfficeOpenXml.Drawing.OleObject
{
    /// <summary>
    /// Types of objects to Embedd
    /// </summary>
    public enum OleObjectType
    {
        /// <summary>
        /// The Default property for most embedded objects.
        /// </summary>
        Default,
        /// <summary>
        /// Use this Ole Object Type for PDF docuemnts for use in Adobe Acrobat. Use Default for other PDF applications.
        /// </summary>
        PDF,
        /// <summary>
        /// Use this Ole Object Type for Libre Office document types.
        /// </summary>
        ODF,
        /// <summary>
        /// Use this Ole Object Type for Microsoft Office document types.
        /// </summary>
        DOC,
    }

    public class ExcelOleObject : ExcelDrawing
    {
        private const string OLE_STREAM_NAME = "\u0001Ole";
        private const string COMPOBJ_STREAM_NAME = "\u0001CompObj";
        private const string OLE10NATIVE_STREAM_NAME = "\u0001Ole10Native";
        private const string CONTENTS_STREAM_NAME = "CONTENTS";
        private const string EMBEDDEDODF_STREAM_NAME = "EmbeddedOdf";

        internal ExcelVmlDrawingBase _vml;
        internal XmlHelper _vmlProp;
        internal OleObjectInternal _oleObject;
        internal CompoundDocument _document;
        internal OleObjectDataStructures _oleDataStructures;
        internal ExcelExternalOleLink _externalLink;
        internal ExcelWorksheet _worksheet;
        internal ZipPackagePart oleObjectPart;
        internal XmlDocument LinkedOleObjectXml;
        internal ZipPackagePart LinkedOleObjectPart;
        internal bool DisplayAsIcon;

        /// <summary>
        /// 
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
        /// <param name="type"></param>
        /// <param name="displayAsIcon"></param>
        /// <param name="iconFilePath"></param>
        /// <param name="parent"></param>
        internal ExcelOleObject(ExcelDrawings drawings, XmlNode node, string filePath, bool linkToFile, OleObjectType type = OleObjectType.Default, bool displayAsIcon = false, string iconFilePath = "", ExcelGroupShape parent = null)
            : base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
        {
            _worksheet = drawings.Worksheet;
            string relId = "";
            string oleObjectNode = "";
            DisplayAsIcon = displayAsIcon;
            if (linkToFile)
            {
                IsExternalLink = true;
                var linkId = CreateLinkToObject(filePath, type);
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
                IsExternalLink = false;
                relId = CreateEmbeddedObject(filePath, type);

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
            int newID = 1;
            var Uri = GetNewUri(_worksheet._package.ZipPackage, "/xl/media/image{0}.emf", ref newID);
            var part = _worksheet._package.ZipPackage.CreatePart(Uri, "image/x-emf", CompressionLevel.None, "emf");
            var rel = _worksheet.Part.CreateRelationship(Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
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
            MemoryStream ms = (MemoryStream)part.GetStream(FileMode.Create, FileAccess.Write);
            ms.Write(image, 0, image.Length);
            var imgRelId = rel.Id;

            //Create drawings xml
            string name = _drawings.GetUniqueDrawingName("Object 1");
            XmlElement spElement = CreateShapeNode();
            spElement.InnerXml = CreateOleObjectDrawingNode(name);
            CreateClientData();
            From.Column = 0;  From.ColumnOff = 0;
            From.Row = 0;     From.RowOff = 0;
            To.Column = 1;    To.ColumnOff = 304800;//171450;
            To.Row = 3;       To.RowOff = 114300;//133350;

            //Create vml
            _vml = drawings.Worksheet.VmlDrawings.AddPicture(this, name, rel.TargetUri);
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
                sb.AppendFormat("<objectPr defaultSize=\"0\" r:id=\"{0}\" dde=\"1\">", imgRelId);
            else
                sb.AppendFormat("<objectPr defaultSize=\"0\" r:id=\"{0}\">", imgRelId);
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
                oleObjectPart = _worksheet._package.ZipPackage.GetPart(oleObj);
                var oleStream = (MemoryStream)oleObjectPart.GetStream(FileMode.Open, FileAccess.Read);
                _document = new CompoundDocument(oleStream);
            }
        }

        private void LoadLinkedObject()
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

        private string CreateEmbeddedObject(string filePath, OleObjectType type)
        {
            string relId = "";
            byte[] fileData = File.ReadAllBytes(filePath);
            string fileType = Path.GetExtension(filePath).ToLower();
            _oleDataStructures = new OleObjectDataStructures();
            _document = new CompoundDocument();
            Guid ClsId = OleObjectGUIDCollection.keyValuePairs["Package"];
            if (type == OleObjectType.PDF) //Only if Acrobat Reader is installed
            {
                //Create Ole structure and add data
                CreateOleObject();
                //Create Ole Data Stream and add to Compound object
                CreateOleDataStream();
                //Create CompObj structure and add data
                CreateCompObjObject("Acrobat Document", "Acrobat.Document.DC");
                //Create CompObj Data Stream and add to Compound object
                CreateCompObjDataStream();
                //Add CONTENT Data Stream
                _oleDataStructures.DataFile = fileData;
                _document.Storage.DataStreams.Add(CONTENTS_STREAM_NAME, new CompoundDocumentItem(CONTENTS_STREAM_NAME, fileData));
                ClsId = new Guid(new byte[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 }); //CHANGE TO PDF GUID?
            }
            else if (type == OleObjectType.ODF) //open office formats if libre office installed
            {
                //Create Ole structure and add data
                CreateOleObject();
                //Create Ole Data Stream and add to Compound object
                CreateOleDataStream();
                //Create CompObj structure and add data
                CreateCompObjObject("OpenDocument Text", "Word.OpenDocumentText.12"); //This has different values depending on if is spreadsheet, presentation or text
                                                                                      //Create CompObj Data Stream and add to Compound object
                CreateCompObjDataStream();
                //Add EmbeddedOdf
                _oleDataStructures.DataFile = fileData;
                _document.Storage.DataStreams.Add(EMBEDDEDODF_STREAM_NAME, new CompoundDocumentItem(EMBEDDEDODF_STREAM_NAME, fileData));
                ClsId = new Guid(new byte[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 }); //CHANGE TO ODF GUID?
            }
            else if (type == OleObjectType.DOC) //ms office format
            {
                //Embedd as is
                string name = "";
                if (fileType == ".docx")
                {
                    name = "Microsoft_Word_Document";
                    CreateCompObjObject("Document", "Document");
                }
                else if (fileType == ".xlsx")
                {
                    name = "Microsoft_Excel_Worksheet";
                    CreateCompObjObject("Worksheet", "Worksheet");
                }
                else if (fileType == ".pptx")
                {
                    name = "Microsoft_PowerPoint_Presentation";
                    CreateCompObjObject("Presentation", "Presentation");
                }
                int newID = 1;
                var Uri = GetNewUri(_worksheet._package.ZipPackage, "/xl/embeddings/" + name + "{0}" + fileType, ref newID);
                var part = _worksheet._package.ZipPackage.CreatePart(Uri, ContentTypes.contentTypeControlProperties); //Change content type or add content type for the doc type?
                var rel = _worksheet.Part.CreateRelationship(Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/embeddings");
                relId = rel.Id;
                MemoryStream ms = (MemoryStream)part.GetStream(FileMode.Create, FileAccess.Write);
                ms.Write(fileData, 0, fileData.Length);
                return relId;
            }
            else if (type == OleObjectType.Default)
            {
                CreateCompObjObject("OLE Package", "Package");
                CreateCompObjDataStream();
                CreateOleNativeObject(fileData, filePath);
                CreateOleNativeDataStream();
                ClsId = OleObjectGUIDCollection.keyValuePairs["Package"];
            }
            if (_document.Storage.DataStreams != null)
            {
                int newID = 1;
                var Uri = GetNewUri(_worksheet._package.ZipPackage, "/xl/embeddings/oleObject{0}.bin", ref newID);
                var part = _worksheet._package.ZipPackage.CreatePart(Uri, ContentTypes.contentTypeOleObject);
                var rel = _worksheet.Part.CreateRelationship(Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/oleObject");
                MemoryStream ms = (MemoryStream)part.GetStream(FileMode.Create, FileAccess.Write);
                _document.RootItem.ClsID = ClsId;
                _document.Save(ms);
                relId = rel.Id;
            }
            return relId;
        }

        private int CreateLinkToObject(string filePath, OleObjectType type)
        {
            var wb = _worksheet.Workbook;
            //create externalLink xml part
            int newID = 1;
            Uri uri = GetNewUri(wb._package.ZipPackage, "/xl/externalLinks/externalLink{0}.xml", ref newID);
            LinkedOleObjectPart = wb._package.ZipPackage.CreatePart(uri, ContentTypes.contentTypeExternalLink);
            var rel = wb.Part.CreateRelationship(uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/externalLink");
            //Create relation to external file
            var fileRel = LinkedOleObjectPart.CreateRelationship("file:///" + filePath, TargetMode.External, ExcelPackage.schemaRelationships + "/oleObject");
            //Create externalLink xml
            //StreamWriter sw = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
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
            LinkedOleObjectXml = new XmlDocument();
            LinkedOleObjectXml.LoadXml(xml.ToString());
            LinkedOleObjectXml.Save(LinkedOleObjectPart.GetStream(FileMode.Create, FileAccess.Write));

            //create/write wb xml external link node
            var er = (XmlElement)wb.CreateNode("d:externalReferences/d:externalReference", false, true);
            er.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);

            //Add the externalLink to externalLink collection
            _externalLink = wb.ExternalLinks[wb.ExternalLinks.GetExternalLink(filePath, fileRel)] as ExcelExternalOleLink; //new ExcelExternalOleLink(wb, new XmlTextReader(LinkedOleObjectPart.GetStream()), LinkedOleObjectPart, er);
            return newID;
        }

        #region WriteBinaries
        private void CreateOleDataStream()
        {
            byte[] oleBytes = BinaryHelper.ConcatenateByteArrays(
                                           BitConverter.GetBytes(_oleDataStructures.Ole.Version),
                                           BitConverter.GetBytes(_oleDataStructures.Ole.Flags),
                                           BitConverter.GetBytes(_oleDataStructures.Ole.LinkUpdateOption),
                                           BitConverter.GetBytes(_oleDataStructures.Ole.Reserved1),
                                           BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStreamSize) );
            if (_oleDataStructures.Ole.ReservedMonikerStreamSize > 0)
            {
                oleBytes = BinaryHelper.ConcatenateByteArrays(oleBytes,
                                        BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStream.ClsId.Data1),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStream.ClsId.Data2),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStream.ClsId.Data3),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStream.ClsId.Data4),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStream.StreamData1),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStream.StreamData2),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStream.StreamData3),
                                        BinaryHelper.GetByteArray(_oleDataStructures.Ole.ReservedMonikerStream.StreamData4, _oleDataStructures.Ole.ReservedMonikerStream.Encoding) );
            }
            if (IsExternalLink)
            {
                oleBytes = BinaryHelper.ConcatenateByteArrays(oleBytes,
                                        BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStreamSize) );
                if (_oleDataStructures.Ole.RelativeSourceMonikerStreamSize > 0)
                {
                    oleBytes = BinaryHelper.ConcatenateByteArrays(oleBytes,
                                            BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStream.ClsId.Data1),
                                            BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStream.ClsId.Data2),
                                            BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStream.ClsId.Data3),
                                            BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStream.ClsId.Data4),
                                            BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStream.StreamData1),
                                            BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStream.StreamData2),
                                            BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStream.StreamData3),
                                            BinaryHelper.GetByteArray(_oleDataStructures.Ole.RelativeSourceMonikerStream.StreamData4, _oleDataStructures.Ole.RelativeSourceMonikerStream.Encoding) );
                }
                oleBytes = BinaryHelper.ConcatenateByteArrays(oleBytes,
                                        BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStreamSize) );
                if (_oleDataStructures.Ole.AbsoluteSourceMonikerStreamSize > 0)
                {
                    oleBytes = BinaryHelper.ConcatenateByteArrays(oleBytes,
                                            BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.ClsId.Data1),
                                            BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.ClsId.Data2),
                                            BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.ClsId.Data3),
                                            BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.ClsId.Data4),
                                            BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.StreamData1),
                                            BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.StreamData2),
                                            BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.StreamData3),
                                            BinaryHelper.GetByteArray(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.StreamData4, _oleDataStructures.Ole.AbsoluteSourceMonikerStream.Encoding) );
                }
                oleBytes = BinaryHelper.ConcatenateByteArrays(oleBytes,
                                        new byte[_oleDataStructures.Ole.ClsIdIndicator],
                                        BitConverter.GetBytes(_oleDataStructures.Ole.ClsId.Data1),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.ClsId.Data2),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.ClsId.Data3),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.ClsId.Data4),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.ReservedDisplayName.Length),
                                        BinaryHelper.GetByteArray(_oleDataStructures.Ole.ReservedDisplayName.String, _oleDataStructures.Ole.ReservedDisplayName.Encoding),
                                        new byte[_oleDataStructures.Ole.Reserved2],
                                        BitConverter.GetBytes(_oleDataStructures.Ole.LocalUpdateTime.dwLowDateTime),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.LocalUpdateTime.dwHighDateTime),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.LocalCheckUpdateTime.dwLowDateTime),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.LocalCheckUpdateTime.dwHighDateTime),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.RemoteUpdateTime.dwLowDateTime),
                                        BitConverter.GetBytes(_oleDataStructures.Ole.RemoteUpdateTime.dwHighDateTime) );
            }
            _document.Storage.DataStreams.Add(OLE_STREAM_NAME, new CompoundDocumentItem(OLE_STREAM_NAME, oleBytes));
        }
        private void CreateCompObjDataStream()
        {
            byte[] compObjBytes = BinaryHelper.ConcatenateByteArrays(
                                               BitConverter.GetBytes(_oleDataStructures.CompObj.Header.Reserved1),
                                               BitConverter.GetBytes(_oleDataStructures.CompObj.Header.Version),
                                               _oleDataStructures.CompObj.Header.Reserved2,
                                               BitConverter.GetBytes(_oleDataStructures.CompObj.AnsiUserType.Length),
                                               BinaryHelper.GetByteArray(_oleDataStructures.CompObj.AnsiUserType.String + "\0", _oleDataStructures.CompObj.AnsiUserType.Encoding),
                                               BitConverter.GetBytes(_oleDataStructures.CompObj.AnsiClipboardFormat.MarkerOrLength),
                                               _oleDataStructures.CompObj.AnsiClipboardFormat.FormatOrAnsiString,
                                               BitConverter.GetBytes(_oleDataStructures.CompObj.Reserved1.Length),
                                               BinaryHelper.GetByteArray(_oleDataStructures.CompObj.Reserved1.String + "\0", _oleDataStructures.CompObj.Reserved1.Encoding),
                                               BitConverter.GetBytes(_oleDataStructures.CompObj.UnicodeMarker),
                                               BitConverter.GetBytes(_oleDataStructures.CompObj.UnicodeUserType.Length),
                                               BinaryHelper.GetByteArray(_oleDataStructures.CompObj.UnicodeUserType.String, _oleDataStructures.CompObj.UnicodeUserType.Encoding),
                                               BitConverter.GetBytes(_oleDataStructures.CompObj.UnicodeClipboardFormat.MarkerOrLength),
                                               _oleDataStructures.CompObj.UnicodeClipboardFormat.FormatOrUnicodeString,
                                               BitConverter.GetBytes(_oleDataStructures.CompObj.Reserved2.Length),
                                               BinaryHelper.GetByteArray(_oleDataStructures.CompObj.Reserved2.String, _oleDataStructures.CompObj.Reserved2.Encoding) );
            _document.Storage.DataStreams.Add(COMPOBJ_STREAM_NAME, new CompoundDocumentItem(COMPOBJ_STREAM_NAME, compObjBytes));
        }
        private void CreateOleNativeDataStream()
        {
            byte[] oleNativeBytes = BinaryHelper.ConcatenateByteArrays(
                                                 BitConverter.GetBytes(_oleDataStructures.OleNative.Header.Size),
                                                 BitConverter.GetBytes(_oleDataStructures.OleNative.Header.Type),
                                                 BinaryHelper.GetByteArray(_oleDataStructures.OleNative.Header.FileName.String + "\0", _oleDataStructures.OleNative.Header.FileName.Encoding),
                                                 BinaryHelper.GetByteArray(_oleDataStructures.OleNative.Header.FilePath.String + "\0", _oleDataStructures.OleNative.Header.FilePath.Encoding),
                                                 BitConverter.GetBytes(_oleDataStructures.OleNative.Header.Reserved1),
                                                 BitConverter.GetBytes(_oleDataStructures.OleNative.Header.TempPath.Length),
                                                 BinaryHelper.GetByteArray(_oleDataStructures.OleNative.Header.TempPath.String + "\0", _oleDataStructures.OleNative.Header.TempPath.Encoding),
                                                 BitConverter.GetBytes(_oleDataStructures.OleNative.NativeDataSize),
                                                 _oleDataStructures.OleNative.NativeData,
                                                 BitConverter.GetBytes(_oleDataStructures.OleNative.Footer.TempPath.Length),
                                                 BinaryHelper.GetByteArray(_oleDataStructures.OleNative.Footer.TempPath.String, _oleDataStructures.OleNative.Footer.TempPath.Encoding),
                                                 BitConverter.GetBytes(_oleDataStructures.OleNative.Footer.FileName.Length),
                                                 BinaryHelper.GetByteArray(_oleDataStructures.OleNative.Footer.FileName.String, _oleDataStructures.OleNative.Footer.FileName.Encoding),
                                                 BitConverter.GetBytes(_oleDataStructures.OleNative.Footer.FilePath.Length),
                                                 BinaryHelper.GetByteArray(_oleDataStructures.OleNative.Footer.FilePath.String, _oleDataStructures.OleNative.Footer.FilePath.Encoding) );
            //Write total size to size.
            var totalsize = BitConverter.GetBytes(oleNativeBytes.Length - 4);
            oleNativeBytes[0] = totalsize[0];
            oleNativeBytes[1] = totalsize[1];
            oleNativeBytes[2] = totalsize[2];
            oleNativeBytes[3] = totalsize[3];
            _document.Storage.DataStreams.Add(OLE10NATIVE_STREAM_NAME, new CompoundDocumentItem(OLE10NATIVE_STREAM_NAME, oleNativeBytes));
        }

        private void CreateOleObject()
        {
            _oleDataStructures.Ole = new OleObjectStream();
            _oleDataStructures.Ole.ReservedMonikerStream = new MonikerStream();
            _oleDataStructures.Ole.ReservedMonikerStream.ClsId = new CLSID();
            byte[] size = BinaryHelper.ConcatenateByteArrays(
                                       BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStream.ClsId.Data1),
                                       BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStream.ClsId.Data2),
                                       BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStream.ClsId.Data3),
                                       BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStream.ClsId.Data4),
                                       BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStream.StreamData1),
                                       BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStream.StreamData2),
                                       BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStream.StreamData3),
                                       BinaryHelper.GetByteArray(_oleDataStructures.Ole.ReservedMonikerStream.StreamData4, _oleDataStructures.Ole.ReservedMonikerStream.Encoding) );
            if (IsExternalLink)
            {
                _oleDataStructures.Ole.ReservedMonikerStreamSize = (UInt32)size.Length;
                _oleDataStructures.Ole.RelativeSourceMonikerStream = new MonikerStream();
                _oleDataStructures.Ole.RelativeSourceMonikerStream.ClsId = new CLSID();
                size = BinaryHelper.ConcatenateByteArrays(
                                    BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStream.ClsId.Data1),
                                    BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStream.ClsId.Data2),
                                    BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStream.ClsId.Data3),
                                    BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStream.ClsId.Data4),
                                    BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStream.StreamData1),
                                    BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStream.StreamData2),
                                    BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStream.StreamData3),
                                    BinaryHelper.GetByteArray(_oleDataStructures.Ole.RelativeSourceMonikerStream.StreamData4, _oleDataStructures.Ole.RelativeSourceMonikerStream.Encoding) );
                _oleDataStructures.Ole.RelativeSourceMonikerStreamSize = (UInt32)size.Length;
                _oleDataStructures.Ole.AbsoluteSourceMonikerStream = new MonikerStream();
                _oleDataStructures.Ole.AbsoluteSourceMonikerStream.ClsId = new CLSID();
                size = BinaryHelper.ConcatenateByteArrays(
                                    BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.ClsId.Data1),
                                    BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.ClsId.Data2),
                                    BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.ClsId.Data3),
                                    BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.ClsId.Data4),
                                    BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.StreamData1),
                                    BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.StreamData2),
                                    BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.StreamData3),
                                    BinaryHelper.GetByteArray(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.StreamData4, _oleDataStructures.Ole.AbsoluteSourceMonikerStream.Encoding) );
                _oleDataStructures.Ole.AbsoluteSourceMonikerStreamSize = (UInt32)size.Length;
                _oleDataStructures.Ole.ClsId = new CLSID();
                _oleDataStructures.Ole.ReservedDisplayName = new LengthPrefixedUnicodeString();
                _oleDataStructures.Ole.LocalUpdateTime = new FILETIME();
                _oleDataStructures.Ole.LocalCheckUpdateTime = new FILETIME();
                _oleDataStructures.Ole.RemoteUpdateTime = new FILETIME();
            }
        }
        private void CreateCompObjObject(string AnsiUserTypeString, string Reserved1String)
        {
            _oleDataStructures.CompObj = new CompObjStream();
            _oleDataStructures.CompObj.Header = new CompObjHeader();
            _oleDataStructures.CompObj.AnsiUserType = new LengthPrefixedAnsiString(AnsiUserTypeString);
            _oleDataStructures.CompObj.AnsiClipboardFormat = new ClipboardFormatOrAnsiString();
            _oleDataStructures.CompObj.Reserved1 = new LengthPrefixedAnsiString(Reserved1String);
            _oleDataStructures.CompObj.UnicodeUserType = new LengthPrefixedUnicodeString();
            _oleDataStructures.CompObj.UnicodeClipboardFormat = new ClipboardFormatOrUnicodeString();
            _oleDataStructures.CompObj.Reserved2 = new LengthPrefixedUnicodeString();
        }

        private void CreateOleNativeObject(byte[] fileData, string filePath)
        {
            _oleDataStructures.OleNative = new OleNativeStream();
            var fileName = Path.GetFileName(filePath);
            var tempLocation = OleObjectDataStructures.GetTempFile(fileName);
            _oleDataStructures.OleNative.Header.FileName.String = fileName;
            _oleDataStructures.OleNative.Header.FilePath.String = filePath;
            _oleDataStructures.OleNative.Header.TempPath = new LengthPrefixedAnsiString(tempLocation);
            _oleDataStructures.OleNative.NativeData = fileData;
            _oleDataStructures.OleNative.NativeDataSize = (uint)fileData.Length;
            _oleDataStructures.OleNative.Footer.TempPath = new LengthPrefixedUnicodeString(tempLocation);
            _oleDataStructures.OleNative.Footer.FileName = new LengthPrefixedUnicodeString(fileName);
            _oleDataStructures.OleNative.Footer.FilePath = new LengthPrefixedUnicodeString(filePath);
        }
        #endregion

        #region ReadBinaries
        private LengthPrefixedUnicodeString ReadLengthPrefixedUnicodeString(BinaryReader br)
        {
            LengthPrefixedUnicodeString LPUniS = new LengthPrefixedUnicodeString();
            LPUniS.Length = br.ReadUInt32();
            LPUniS.String = BinaryHelper.GetString(br, LPUniS.Length * 2, LPUniS.Encoding);
            return LPUniS;
        }
        private LengthPrefixedAnsiString ReadLengthPrefixedAnsiString(BinaryReader br)
        {
            LengthPrefixedAnsiString LPAnsiS = new LengthPrefixedAnsiString();
            LPAnsiS.Length = br.ReadUInt32();
            LPAnsiS.String = BinaryHelper.GetString(br, LPAnsiS.Length, LPAnsiS.Encoding).Trim('\0');
            return LPAnsiS;
        }
        private LengthPrefixedAnsiString ReadUntilNullTerminator(BinaryReader br)
        {
            LengthPrefixedAnsiString LPAnsiS = new LengthPrefixedAnsiString();
            List<byte> bytes = new List<byte>();
            byte b;
            while ((b = br.ReadByte()) != 0x00)
            {
                bytes.Add(b);
            }
            LPAnsiS.String = BinaryHelper.GetString(bytes.ToArray(), Encoding.ASCII);
            return LPAnsiS;
        }
        private ClipboardFormatOrUnicodeString ReadClipboardFormatOrUnicodeString(BinaryReader br)
        {
            ClipboardFormatOrUnicodeString CFOUS = new ClipboardFormatOrUnicodeString();
            CFOUS.MarkerOrLength = br.ReadUInt32();
            if (CFOUS.MarkerOrLength > 0x00000190 || CFOUS.MarkerOrLength == 0x00000000)
            {
                CFOUS.FormatOrUnicodeString = null;
            }
            else if (CFOUS.MarkerOrLength == 0xFFFFFFFF || CFOUS.MarkerOrLength == 0xFFFFFFFE)
            {
                CFOUS.FormatOrUnicodeString = br.ReadBytes(4);
            }
            else
            {
                CFOUS.FormatOrUnicodeString = br.ReadBytes((int)CFOUS.MarkerOrLength); //This is a string
            }
            return CFOUS;
        }
        private ClipboardFormatOrAnsiString ReadClipboardFormatOrAnsiString(BinaryReader br)
        {
            ClipboardFormatOrAnsiString CFOAS = new ClipboardFormatOrAnsiString();
            CFOAS.MarkerOrLength = br.ReadUInt32();
            if (CFOAS.MarkerOrLength > 0x00000190 || CFOAS.MarkerOrLength == 0x00000000)
            {
                CFOAS.FormatOrAnsiString = null;
            }
            else if (CFOAS.MarkerOrLength == 0xFFFFFFFF || CFOAS.MarkerOrLength == 0xFFFFFFFE)
            {
                CFOAS.FormatOrAnsiString = br.ReadBytes(4);
            }
            else
            {
                CFOAS.FormatOrAnsiString = br.ReadBytes((int)CFOAS.MarkerOrLength); //This is a string
            }
            return CFOAS;
        }
        private CLSID ReadCLSID(BinaryReader br)
        {
            CLSID CLSID = new CLSID();
            CLSID.Data1 = br.ReadUInt32();
            CLSID.Data2 = br.ReadUInt16();
            CLSID.Data3 = br.ReadUInt16();
            CLSID.Data4 = br.ReadUInt64();
            return CLSID;
        }
        private MonikerStream ReadMONIKERSTREAM(BinaryReader br, uint size)
        {
            MonikerStream monikerStream = new MonikerStream();
            monikerStream.ClsId = ReadCLSID(br);
            monikerStream.StreamData1 = br.ReadUInt32();
            monikerStream.StreamData2 = br.ReadUInt16();
            monikerStream.StreamData3 = br.ReadUInt32();
            monikerStream.StreamData4 = BinaryHelper.GetString(br, monikerStream.StreamData3, Encoding.ASCII);
            return monikerStream;
        }
        private FILETIME ReadFILETIME(BinaryReader br)
        {
            FILETIME FILETIME = new FILETIME();
            FILETIME.dwLowDateTime = br.ReadUInt32();
            FILETIME.dwHighDateTime = br.ReadUInt32();
            return FILETIME;
        }
        private void ReadOleStream(byte[] oleBytes)
        {
            using (var ms = RecyclableMemory.GetStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);
                _oleDataStructures.Ole.Version = br.ReadUInt32();
                _oleDataStructures.Ole.Flags = br.ReadUInt32();
                _oleDataStructures.Ole.LinkUpdateOption = br.ReadUInt32();
                _oleDataStructures.Ole.Reserved1 = br.ReadUInt32();
                _oleDataStructures.Ole.ReservedMonikerStreamSize = br.ReadUInt32();
                if (_oleDataStructures.Ole.ReservedMonikerStreamSize != 0)
                    _oleDataStructures.Ole.ReservedMonikerStream = ReadMONIKERSTREAM(br, _oleDataStructures.Ole.ReservedMonikerStreamSize - 4);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.Ole.RelativeSourceMonikerStreamSize = br.ReadUInt32();
                if (_oleDataStructures.Ole.RelativeSourceMonikerStreamSize != 0)
                    _oleDataStructures.Ole.RelativeSourceMonikerStream = ReadMONIKERSTREAM(br, _oleDataStructures.Ole.RelativeSourceMonikerStreamSize - 4);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.Ole.AbsoluteSourceMonikerStreamSize = br.ReadUInt32();
                if (_oleDataStructures.Ole.AbsoluteSourceMonikerStreamSize != 0)
                    _oleDataStructures.Ole.AbsoluteSourceMonikerStream = ReadMONIKERSTREAM(br, _oleDataStructures.Ole.AbsoluteSourceMonikerStreamSize - 4);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.Ole.ClsIdIndicator = br.ReadUInt32();
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.Ole.ClsId = ReadCLSID(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.Ole.ReservedDisplayName = ReadLengthPrefixedUnicodeString(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.Ole.Reserved2 = br.ReadUInt32();
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.Ole.LocalUpdateTime = ReadFILETIME(br);
                _oleDataStructures.Ole.LocalCheckUpdateTime = ReadFILETIME(br);
                _oleDataStructures.Ole.RemoteUpdateTime = ReadFILETIME(br);
            }
        }
        private CompObjHeader ReadCompObjHeader(BinaryReader br)
        {
            CompObjHeader header = new CompObjHeader();
            header.Reserved1 = br.ReadUInt32();
            header.Version = br.ReadUInt32();
            header.Reserved2 = br.ReadBytes(20);
            return header;
        }
        private void ReadCompObjStream(byte[] oleBytes)
        {
            using (var ms = RecyclableMemory.GetStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);
                _oleDataStructures.CompObj.Header = ReadCompObjHeader(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.CompObj.AnsiUserType = ReadLengthPrefixedAnsiString(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.CompObj.AnsiClipboardFormat = ReadClipboardFormatOrAnsiString(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.CompObj.Reserved1 = ReadLengthPrefixedAnsiString(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.CompObj.UnicodeMarker = br.ReadUInt32();
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.CompObj.UnicodeUserType = ReadLengthPrefixedUnicodeString(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.CompObj.UnicodeClipboardFormat = ReadClipboardFormatOrUnicodeString(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.CompObj.Reserved2 = ReadLengthPrefixedUnicodeString(br);
            }
        }
        private OleNativeHeader ReadOleNativeHeader(BinaryReader br)
        {
            OleNativeHeader header = new OleNativeHeader();
            header.Size = br.ReadUInt32();
            header.Type = br.ReadUInt16();
            header.FileName = ReadUntilNullTerminator(br);
            header.FilePath = ReadUntilNullTerminator(br);
            header.Reserved1 = br.ReadUInt32();
            header.TempPath = ReadLengthPrefixedAnsiString(br);
            return header;
        }
        private OleNativeFooter ReadOleNativeFooter(BinaryReader br)
        {
            OleNativeFooter footer = new OleNativeFooter();
            footer.TempPath = ReadLengthPrefixedUnicodeString(br);
            footer.FileName = ReadLengthPrefixedUnicodeString(br);
            footer.FilePath = ReadLengthPrefixedUnicodeString(br);
            return footer;
        }
        private void ReadOleNative(byte[] oleBytes)
        {
            using (var ms = RecyclableMemory.GetStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);
                _oleDataStructures.OleNative.Header = ReadOleNativeHeader(br);
                _oleDataStructures.OleNative.NativeDataSize = br.ReadUInt32();
                _oleDataStructures.OleNative.NativeData = br.ReadBytes((int)_oleDataStructures.OleNative.NativeDataSize);
                _oleDataStructures.OleNative.Footer = ReadOleNativeFooter(br);
            }
        }
        #endregion

        #region Export
        internal void ExportOleObjectData(string ExportPath)
        {
            _oleDataStructures = new OleObjectDataStructures();
            if (_document.Storage.DataStreams.ContainsKey(OLE10NATIVE_STREAM_NAME))
            {
                _oleDataStructures.OleNative = new OleObjectDataStructures.OleNativeStream();
                ReadOleNative(_document.Storage.DataStreams[OLE10NATIVE_STREAM_NAME].Stream);
            }
            if (_document.Storage.DataStreams.ContainsKey(OLE_STREAM_NAME))
            {
                _oleDataStructures.Ole = new OleObjectDataStructures.OleObjectStream();
                ReadOleStream(_document.Storage.DataStreams[OLE_STREAM_NAME].Stream);
            }
            if (_document.Storage.DataStreams.ContainsKey(COMPOBJ_STREAM_NAME))
            {
                _oleDataStructures.CompObj = new OleObjectDataStructures.CompObjStream();
                ReadCompObjStream(_document.Storage.DataStreams[COMPOBJ_STREAM_NAME].Stream);
            }
            using var p = new ExcelPackage(ExportPath);
            OleObjectDataStreamsExport.ExportOleNative(_worksheet._package.File.Name, oleObjectPart.Entry.FileName, p, _oleDataStructures);
            OleObjectDataStreamsExport.ExportOle(_worksheet._package.File.Name, oleObjectPart.Entry.FileName, p, _oleDataStructures, IsExternalLink);
            OleObjectDataStreamsExport.ExportCompObj(_worksheet._package.File.Name, oleObjectPart.Entry.FileName, p, _oleDataStructures);
            p.Save();
        }
        #endregion
    }
}

/*
 * TODO:
 * DELETE OleObject
 * Copy OleObject
 * 
 * Skapa default värden för aString och Resereved1String i CompObj
 * Funktion för att sätta StreamData4 i Ole som är worksheetName!ObjectName
 *
 *
 * user specidified aString och Reserved1String
 * user specified Image
 *
 *
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
 *      Olika mängd filer, de viktiga är Ole, OleNative, CompObj, samt potentiellt en fil som är själva filen(CONTENT för t ex en pdf), och OlePresXXX
 *      Ole
 *          Existerar -> Skriv ny data till filen
 *          Existerar inte -> Skapa filen om vi inte ska skapa en OleNative
 *      CompObj är de vi ska skriva data till. När vi sparar. Dessa får vi skapa när vi embeddar ett objekt som har dessa filer.
 *          Exsisterar -> Skriv data till filen
 *          Existerar inte -> Skapa filen
 *      OleNative
 *          Existerar -> Ingen skrivning till filen
 *          Existerar inte -> Skapa filen om den behövs om det inte ska skapas någon Ole-fil
 *      OlePres
 *          Existerar -> Ingen skrivning till filen
 *          Existerar inte -> Skapa aldrig filen.
 *      CONTENT
 *          Själva PDF filen i ett compound objekt. Måste exsistera
 *      EmbeddedOdf
 *          så kallas själva filen för de öppna office typerna
 *
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
 *
 * I Excel
 *  I Create New fliken så vill excel skapa nytt dokument och göra editering direkt.
 *  Skapar man ett package så blir det en oleNative oavsett verkar det som.
 *  Create from file kikar på filändelsen verkar det som och skapar filen baserat på det.
 *
 *
 */

/*
 * Add file as oleObject
 * We insert a path to the file
 * We can add guid for application to open?   //Högst otroligt att vi använder denna. Lär nog använda package, men att kunna specifiera pdf eller odf format som optionals kan vara en lösning, notera att för pdf så tar adobe reader över helt även om filen är ett package.
 * we check what file type it is.
 */
