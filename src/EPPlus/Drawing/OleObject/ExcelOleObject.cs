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

                /*
                Skapa relation till .bin filen. Denna relation går från worksheet till embeddings/oleObjectX.bin.
                detta gör vi genom att skapa en uri och en part som sedan ger oss relations id.
                Vi använder GetNewUri
                Sedan gör vi CreatePart? Vi måste nog uppdatera ContentTypes så den har en oleObject typ.
                Sedan skapar vi relationen som vi sedan har när vi skriver xml.

                Sedan måste vi skapa .bin filen. Detta görs genom att använda CompoundDokument på något vis. Problemet här är att
                just nu har vi inget bra sätt att ge ett namn och placera vår compound dokument i embeddings mappen?

                I save HandleSaveForIndividualDrawings måste vi uppdatera för support för oleObjet?
                Är det något mer i save som måste göras?

                */


                int newID = 1;
                var Uri = GetNewUri(_worksheet._package.ZipPackage, "/xl/embeddings/oleObject{0}.bin", ref newID);
                var part = _worksheet._package.ZipPackage.CreatePart(Uri, ContentTypes.contentTypeControlProperties);
                var rel = _worksheet.Part.CreateRelationship(Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/embeddings");

                MemoryStream ms = (MemoryStream)part.GetStream(FileMode.Create, FileAccess.Write);
                byte[] data = File.ReadAllBytes(filepath);
                _document = new CompoundDocument();
                _document.Storage.DataStreams.Add("\u0001Ole",CreateOleStream());
                _document.Storage.DataStreams.Add("\u0001CompObj", CreateCompObjStream());
                _document.Storage.DataStreams.Add("CONTENTS", data);
                _document.Save(ms);
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

        private byte[] CreateCompObjStream()
        {
            throw new NotImplementedException();
        }

        private byte[] CreateOleStream()
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                BinaryWriter bw = new BinaryWriter(ms);

                /****** PROJECTINFORMATION Record ******/
                bw.Write((uint)0x02000001);        //Version
                bw.Write((uint)0x00000000);          //Flags

                ms.Flush();
                return ms.ToArray();
            }
        }

        private byte[] ReadClipboardFormatOrAnsiString(BinaryReader br)
        {
            var MarkerOrLength = br.ReadUInt32();
            byte[] FormatOrAnsiString = null;
            if (MarkerOrLength > 0x00000190 || MarkerOrLength == 0x00000000)
            {
                return new byte[] { }; //error
            }
            else if (MarkerOrLength == 0xFFFFFFFF || MarkerOrLength == 0xFFFFFFFE)
            {
                FormatOrAnsiString = br.ReadBytes(4);
            }
            else
            {
                FormatOrAnsiString = br.ReadBytes((int)MarkerOrLength); //This is a string
            }
            return FormatOrAnsiString;
        }

        private byte[] ReadClipboardFormatOrUnicodeString(BinaryReader br)
        {
            var MarkerOrLength = br.ReadUInt32();
            byte[] FormatOrUnicodeString = null;
            if (MarkerOrLength > 0x00000190 || MarkerOrLength == 0x00000000)
            {
                return new byte[] { }; //error
            }
            else if (MarkerOrLength == 0xFFFFFFFF || MarkerOrLength == 0xFFFFFFFE)
            {
                FormatOrUnicodeString = br.ReadBytes(4);
            }
            else
            {
                FormatOrUnicodeString = br.ReadBytes((int)MarkerOrLength); //This is a string
            }
            return FormatOrUnicodeString;
        }

        private void ReadTOCENTRY(BinaryReader br)
        {
            ReadClipboardFormatOrAnsiString(br); //AnsiClipboardFormat
            var TargetDeviceSize = br.ReadUInt32();
            var Aspect = br.ReadUInt32();
            var Lindex = br.ReadUInt32();
            var Tymed = br.ReadUInt32();
            var Reserved1 = br.ReadBytes(12);
            var Advf = br.ReadUInt32();
            var Reserved2 = br.ReadUInt32();
            ReadDVTARGETDEVICE(br); //TargetDevice
        }

        private void ReadDEVMODEA(BinaryReader br)
        {
            var dmDeviceName = br.ReadBytes(32);
            var dmFormName = br.ReadBytes(32);
            var dmSpecVersion = br.ReadUInt16();
            var dmDriverVersion = br.ReadUInt16();
            var dmSize = br.ReadUInt16();
            var dmDriverExtra = br.ReadUInt16();
            var dmFields = br.ReadUInt32();
            var dmOrientation = br.ReadUInt16();
            var dmPaperSize = br.ReadUInt16();
            var dmPaperLength = br.ReadUInt16();
            var dmPaperWidth = br.ReadUInt16();
            var dmScale = br.ReadUInt16();
            var dmCopies = br.ReadUInt16();
            var dmDefaultSource = br.ReadUInt16();
            var dmPrintQuality = br.ReadUInt16();
            var dmColor = br.ReadUInt16();
            var dmDuplex = br.ReadUInt16();
            var dmYResolution = br.ReadUInt16();
            var dmTTOption = br.ReadUInt16();
            var dmCollate = br.ReadUInt16();
            var reserved0 = br.ReadUInt32();
            var reserved1 = br.ReadUInt32();
            var reserved2 = br.ReadUInt32();
            var reserved3 = br.ReadUInt32();
            var dmNup = br.ReadUInt32();
            var reserved4 = br.ReadUInt32();
            var dmICMMethod = br.ReadUInt32();
            var dmICMIntent = br.ReadUInt32();
            var dmMediaType = br.ReadUInt32();
            var dmDitherType = br.ReadUInt32();
            var reserved5 = br.ReadUInt32();
            var reserved6 = br.ReadUInt32();
            var reserved7 = br.ReadUInt32();
            var reserved8 = br.ReadUInt32();
        }

        private void ReadDVTARGETDEVICE(BinaryReader br)
        {
            var DriverNameOffSet = br.ReadUInt16();
            var DeviceNameOffSet = br.ReadUInt16();
            var PortNameOffSet = br.ReadUInt16();
            var ExtDevModeOffSet = br.ReadUInt16();

            string DriverName = "";
            var DriverNameLength = ExtDevModeOffSet - PortNameOffSet - DeviceNameOffSet - DriverNameOffSet;
            if (DriverNameOffSet != 0)
                DriverName = BinaryHelper.GetString(br, (uint)DriverNameLength, Encoding.ASCII);

            string DeviceName = "";
            var DeviceNameLength = ExtDevModeOffSet - PortNameOffSet - DeviceNameOffSet;
            if (DeviceNameOffSet != 0)
                DeviceName = BinaryHelper.GetString(br, (uint)DeviceNameLength, Encoding.ASCII);

            string PortName = "";
            var PortNameLength = ExtDevModeOffSet - PortNameOffSet;
            if (PortNameOffSet != 0)
                PortName = BinaryHelper.GetString(br, (uint)PortNameLength, Encoding.ASCII);

            if (ExtDevModeOffSet != 0)
                ReadDEVMODEA(br); //ExtDevMode
        }
        private void ReadMONIKERSTREAM(BinaryReader br, uint size)
        {
            var ClsId = br.ReadBytes(16);
            var StreamData1 = br.ReadUInt32();
            var StreamData2 = br.ReadUInt16();
            var StreamData3 = br.ReadUInt32();
            var StreamData4 = BinaryHelper.GetString(br, StreamData3, Encoding.ASCII);
        }

        private void ReadCLSID(BinaryReader br)
        {
            var Data1 = br.ReadUInt32();
            var Data2 = br.ReadUInt16();
            var Data3 = br.ReadUInt16();
            var Data4 = br.ReadUInt64();
        }

        private void ReadLengthPrefixedUnicodeString(BinaryReader br)
        {
            var Length = br.ReadUInt32();
            var uString = BinaryHelper.GetString(br, Length, Encoding.Unicode);
        }

        private void ReadLengthPrefixedAnsiString(BinaryReader br)
        {
            var Length = br.ReadUInt32();
            var aString = BinaryHelper.GetString(br, Length, Encoding.ASCII);
        }

        private void ReadFILETIME(BinaryReader br)
        {
            var dwLowDateTime = br.ReadUInt32();
            var dwHighDateTime = br.ReadUInt32();
        }

        private void ReadCompObjHeader(BinaryReader br)
        {
            var Reserved1 = br.ReadUInt32();
            var Version = br.ReadUInt32();
            var Reserved2 = br.ReadBytes(20);
        }

        private void ReadCompObjStream(byte[] oleBytes)
        {
            using (var ms = RecyclableMemory.GetStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);

                ReadCompObjHeader(br); //Header

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadLengthPrefixedAnsiString(br); //AnsiUserType

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadClipboardFormatOrAnsiString(br); //AnsiClipboardFormat 

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                //Reserved1 should be a LengthPrefixedUnicodeString
                var Reserved1Length = br.ReadUInt32();
                if (Reserved1Length == 0 || Reserved1Length > 0x00000028)
                {
                    return;
                }
                var Reserved1String = BinaryHelper.GetString(br, Reserved1Length, Encoding.ASCII);
                if (string.IsNullOrEmpty(Reserved1String))
                {
                    return;
                }
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                var UnicodeMarker = br.ReadUInt32();
                if(UnicodeMarker != 0x71B239F4)
                {
                    return;
                }
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadLengthPrefixedUnicodeString(br); //UnicodeUserType
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadClipboardFormatOrUnicodeString(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadLengthPrefixedUnicodeString(br); //Reserved2
            }
        }

        private void ReadOleStream(byte[] oleBytes)
        {
            using (var ms = RecyclableMemory.GetStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);
                var Version = br.ReadUInt32();
                var Flags = br.ReadUInt32();
                var LinkUpdateOption = br.ReadUInt32();
                var Reserved1 = br.ReadUInt32();
                var ReservedMonikerStreamSize = br.ReadUInt32() - 4;
                if (ReservedMonikerStreamSize != 0)
                {
                    ReadMONIKERSTREAM(br, ReservedMonikerStreamSize);
                }

                if( br.BaseStream.Position >= br.BaseStream.Length )
                    return;

                var RelativeSourceMonikerStreamSize = br.ReadUInt32() -4;
                if (RelativeSourceMonikerStreamSize != 0)
                {
                    ReadMONIKERSTREAM(br, RelativeSourceMonikerStreamSize);
                }

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                var AbsoluteSourceMonikerStreamSize = br.ReadUInt32() - 4;
                if (AbsoluteSourceMonikerStreamSize != 0)
                {
                    ReadMONIKERSTREAM(br, AbsoluteSourceMonikerStreamSize);
                }

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                var ClsidIndicator = br.ReadUInt32();

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadCLSID(br); //Clsid

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadLengthPrefixedUnicodeString(br); //ReservedDisplayName

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                var Reserved2 = br.ReadUInt32();

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadFILETIME(br); //LocalUpdateTime
                ReadFILETIME(br); //LocalCheckUpdateTime
                ReadFILETIME(br); //RemoteUpdateTime
            }
        }

        private void ReadOleNative(byte[] oleBytes)
        {
            using (var ms = RecyclableMemory.GetStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);
                var NativeDataSize = br.ReadUInt32();
                var NativeData = br.ReadBytes((int)NativeDataSize);
            }
        }

        private void ReadOlePres(byte[] oleBytes)
        {
            using (var ms = new MemoryStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);
                var AnsiClipboardFormatFormatOrAnsiString = ReadClipboardFormatOrAnsiString(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;
                var TargetDeviceSize = br.ReadUInt32();
                if (TargetDeviceSize >= 0x00000004)
                {
                    ReadDVTARGETDEVICE(br); //TargetDevice
                }
                var Aspect = br.ReadUInt32();
                var Lindex = br.ReadUInt32();
                var Advf = br.ReadUInt32();
                var Reserved1 = br.ReadUInt32();
                var Width = br.ReadUInt32();
                var Height = br.ReadUInt32();
                var Size = br.ReadUInt32();
                var Data = br.ReadBytes((int)Size);
                byte[] Reserved2;
                if (BitConverter.ToInt32( AnsiClipboardFormatFormatOrAnsiString, 0 ) == 0x00000003)
                    Reserved2 = br.ReadBytes(18);
                var TocSignature = br.ReadUInt32();
                var TocCount = br.ReadUInt32();
                if (TocSignature == 0x494E414 || TocCount == 0)
                    return;
                for (int i = 0; i < TocCount; i++)
                {
                    ReadTOCENTRY(br);
                }
            }
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
            if (_document.Storage.DataStreams.ContainsKey("\u0001Ole10Native"))
                ReadOleNative(_document.Storage.DataStreams["\u0001Ole10Native"]);
            if (_document.Storage.DataStreams.ContainsKey("\u0001Ole"))
                ReadOleStream(_document.Storage.DataStreams["\u0001Ole"]);
            if (_document.Storage.DataStreams.ContainsKey("\u0001CompObj"))
                ReadCompObjStream(_document.Storage.DataStreams["\u0001CompObj"]);
            for(int i = 0; i <= 999; i++)
            {
                string olePres = "\u0010OlePres" + i.ToString("D3");
                if (_document.Storage.DataStreams.ContainsKey(olePres))
                    ReadCompObjStream(_document.Storage.DataStreams[olePres]);
            }
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