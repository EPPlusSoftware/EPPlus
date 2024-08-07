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
using OfficeOpenXml.Encryption;
using static OfficeOpenXml.Drawing.OleObject.OleObjectDataStreams;
using OfficeOpenXml.Core.Worksheet.XmlWriter;


namespace OfficeOpenXml.Drawing.OleObject
{
    public class ExcelOleObject : ExcelDrawing
    {
        internal ExcelVmlDrawingBase _vml;
        internal XmlHelper _vmlProp;
        internal OleObjectInternal _oleObject;
        internal CompoundDocument _document;
        internal OleObjectDataStreams _oleDataStreams;
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

        private string CreateOleObjectDrawingNode()
        {
            StringBuilder xml = new StringBuilder();
            xml.Append($"<xdr:nvSpPr><xdr:cNvPr hidden=\"1\" name=\"\" id=\"{_id}\"><a:extLst><a:ext uri=\"{{63B3BB69-23CF-44E3-9099-C40C66FF867C}}\"><a14:compatExt spid=\"_x0000_s{_id}\"/></a:ext><a:ext uri=\"{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}\"><a16:creationId id=\"{{00000000-0008-0000-0000-000001040000}}\" xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr>");
            xml.Append($"<xdr:spPr bwMode=\"auto\"><a:xfrm><a:off y=\"0\" x=\"0\"/><a:ext cy=\"0\" cx=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
            xml.Append($"<a:solidFill><a:srgbClr val=\"FFFFFF\" mc:Ignorable=\"a14\" a14:legacySpreadsheetColorIndex=\"65\"/></a:solidFill><a:ln w=\"9525\"><a:solidFill><a:srgbClr val=\"000000\" mc:Ignorable=\"a14\" a14:legacySpreadsheetColorIndex=\"64\"/></a:solidFill><a:miter lim=\"800000\"/><a:headEnd/><a:tailEnd/></a:ln></xdr:spPr>");
            return xml.ToString();
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

        private void ExportClipboardFormatOrAnsiString(ExcelWorksheet ws, ref int ci, OleObjectDataStreams.ClipboardFormatOrAnsiString CFOAS)
        {
            ws.Cells[2, ci++].Value = CFOAS.MarkerOrLength;
            ws.Cells[2, ci++].Value = CFOAS.FormatOrAnsiString;
        }

        private OleObjectDataStreams.ClipboardFormatOrAnsiString ReadClipboardFormatOrAnsiString(BinaryReader br)
        {
            OleObjectDataStreams.ClipboardFormatOrAnsiString CFOAS = new OleObjectDataStreams.ClipboardFormatOrAnsiString();
            CFOAS.MarkerOrLength = br.ReadUInt32();
            if (CFOAS.MarkerOrLength > 0x00000190 || CFOAS.MarkerOrLength == 0x00000000)
            {
                CFOAS.FormatOrAnsiString = null;
            }
            else if ( CFOAS.MarkerOrLength == 0xFFFFFFFF || CFOAS.MarkerOrLength == 0xFFFFFFFE)
            {
                CFOAS.FormatOrAnsiString = br.ReadBytes(4);
            }
            else
            {
                CFOAS.FormatOrAnsiString = br.ReadBytes((int)CFOAS.MarkerOrLength); //This is a string
            }
            return CFOAS;
        }

        private void ExportClipboardFormatOrUnicodeString(ExcelWorksheet ws, ref int ci, OleObjectDataStreams.ClipboardFormatOrUnicodeString CFOUS)
        {
            ws.Cells[2, ci++].Value = CFOUS.MarkerOrLength;
            ws.Cells[2, ci++].Value = CFOUS.FormatOrUnicodeString;
        }

        private OleObjectDataStreams.ClipboardFormatOrUnicodeString ReadClipboardFormatOrUnicodeString(BinaryReader br)
        {
            OleObjectDataStreams.ClipboardFormatOrUnicodeString CFOUS = new OleObjectDataStreams.ClipboardFormatOrUnicodeString();
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

        //private void ReadTOCENTRY(BinaryReader br, ExcelWorksheet ws, ref int ci)
        //{
        //    ReadClipboardFormatOrAnsiString(br, ws, ref ci); //AnsiClipboardFormat
        //    var TargetDeviceSize = br.ReadUInt32();
        //    var Aspect = br.ReadUInt32();
        //    var Lindex = br.ReadUInt32();
        //    var Tymed = br.ReadUInt32();
        //    var Reserved1 = br.ReadBytes(12);
        //    var Advf = br.ReadUInt32();
        //    var Reserved2 = br.ReadUInt32();

        //    ws.Cells[2, ci++].Value = TargetDeviceSize;
        //    ws.Cells[2, ci++].Value = Aspect;
        //    ws.Cells[2, ci++].Value = Lindex;
        //    ws.Cells[2, ci++].Value = Tymed;
        //    ws.Cells[2, ci++].Value = Reserved1;
        //    ws.Cells[2, ci++].Value = Advf;
        //    ws.Cells[2, ci++].Value = Reserved2;

        //    ReadDVTARGETDEVICE(br, TargetDeviceSize, ws, ref ci); //TargetDevice
        //}

        private void ReadDEVMODEA(BinaryReader br, ExcelWorksheet ws, ref int ci)
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

            ws.Cells[2, ci++].Value = dmDeviceName;
            ws.Cells[2, ci++].Value = dmFormName;
            ws.Cells[2, ci++].Value = dmSpecVersion;
            ws.Cells[2, ci++].Value = dmDriverVersion;
            ws.Cells[2, ci++].Value = dmSize;
            ws.Cells[2, ci++].Value = dmDriverExtra;
            ws.Cells[2, ci++].Value = dmFields;
            ws.Cells[2, ci++].Value = dmOrientation;
            ws.Cells[2, ci++].Value = dmPaperSize;
            ws.Cells[2, ci++].Value = dmPaperLength;
            ws.Cells[2, ci++].Value = dmPaperWidth;
            ws.Cells[2, ci++].Value = dmScale;
            ws.Cells[2, ci++].Value = dmCopies;
            ws.Cells[2, ci++].Value = dmDefaultSource;
            ws.Cells[2, ci++].Value = dmPrintQuality;
            ws.Cells[2, ci++].Value = dmColor;
            ws.Cells[2, ci++].Value = dmDuplex;
            ws.Cells[2, ci++].Value = dmYResolution;
            ws.Cells[2, ci++].Value = dmTTOption;
            ws.Cells[2, ci++].Value = dmCollate;
            ws.Cells[2, ci++].Value = reserved0;
            ws.Cells[2, ci++].Value = reserved1;
            ws.Cells[2, ci++].Value = reserved2;
            ws.Cells[2, ci++].Value = reserved3;
            ws.Cells[2, ci++].Value = dmNup;
            ws.Cells[2, ci++].Value = reserved4;
            ws.Cells[2, ci++].Value = dmICMMethod;
            ws.Cells[2, ci++].Value = dmICMIntent;
            ws.Cells[2, ci++].Value = dmMediaType;
            ws.Cells[2, ci++].Value = dmDitherType;
            ws.Cells[2, ci++].Value = reserved5;
            ws.Cells[2, ci++].Value = reserved6;
            ws.Cells[2, ci++].Value = reserved7;
            ws.Cells[2, ci++].Value = reserved8;
        }


        static ushort MinOffset(ushort[] offsets, ushort currentOffset)
        {
            ushort minOffset = ushort.MaxValue;
            foreach (ushort offset in offsets)
            {
                if (offset > currentOffset && offset < minOffset)
                {
                    minOffset = offset;
                }
            }
            return minOffset;
        }

        private void ReadDVTARGETDEVICE(BinaryReader br, uint size, ExcelWorksheet ws, ref int ci)
        {
            var DriverNameOffSet = br.ReadUInt16();
            var DeviceNameOffSet = br.ReadUInt16();
            var PortNameOffSet = br.ReadUInt16();
            var ExtDevModeOffSet = br.ReadUInt16();

            ws.Cells[2, ci++].Value = DriverNameOffSet;
            ws.Cells[2, ci++].Value = DeviceNameOffSet;
            ws.Cells[2, ci++].Value = PortNameOffSet;
            ws.Cells[2, ci++].Value = ExtDevModeOffSet;

            string DriverName = "";
            if (DriverNameOffSet != 0)
            {
                ushort nextOffset = MinOffset(new ushort[] { DeviceNameOffSet, PortNameOffSet, ExtDevModeOffSet, (ushort)size }, DriverNameOffSet);
                var DriverNameLength = nextOffset - DriverNameOffSet;
                DriverName = BinaryHelper.GetString(br, (uint)DriverNameLength, Encoding.ASCII);
            }

            ws.Cells[2, ci++].Value = DriverName;

            string DeviceName = "";

            if (DeviceNameOffSet != 0)
            {
                ushort nextOffset = MinOffset(new ushort[] { DriverNameOffSet, PortNameOffSet, ExtDevModeOffSet, (ushort)size }, DeviceNameOffSet);
                var DeviceNameLength = nextOffset - DeviceNameOffSet;
                DeviceName = BinaryHelper.GetString(br, (uint)DeviceNameLength, Encoding.ASCII);
            }

            ws.Cells[2, ci++].Value = DeviceName;

            string PortName = "";
            if (PortNameOffSet != 0)
            {
                ushort nextOffset = MinOffset(new ushort[] { DriverNameOffSet, DeviceNameOffSet, ExtDevModeOffSet, (ushort)size }, PortNameOffSet);
                var PortNameLength = nextOffset - PortNameOffSet;
                PortName = BinaryHelper.GetString(br, (uint)PortNameLength, Encoding.ASCII);
            }

            ws.Cells[2, ci++].Value = PortName;

            if (ExtDevModeOffSet != 0)
                ReadDEVMODEA(br, ws, ref ci); //ExtDevMode
            else
                ci += 34;
        }
        private OleObjectDataStreams.MonikerStream ReadMONIKERSTREAM(BinaryReader br, uint size)
        {
            OleObjectDataStreams.MonikerStream monikerStream = new OleObjectDataStreams.MonikerStream();
            monikerStream.ClsId = ReadCLSID(br);
            monikerStream.StreamData1 = br.ReadUInt32();
            monikerStream.StreamData2 = br.ReadUInt16();
            monikerStream.StreamData3 = br.ReadUInt32();
            monikerStream.StreamData4 = BinaryHelper.GetString(br, monikerStream.StreamData3, Encoding.Unicode);
            return monikerStream;
        }

        private OleObjectDataStreams.CLSID ReadCLSID(BinaryReader br)
        {
            OleObjectDataStreams.CLSID CLSID = new OleObjectDataStreams.CLSID();
            CLSID.Data1 = br.ReadUInt32();
            CLSID.Data2 = br.ReadUInt16();
            CLSID.Data3 = br.ReadUInt16();
            CLSID.Data4 = br.ReadUInt64();
            return CLSID;
        }

        private OleObjectDataStreams.LengthPrefixedUnicodeString ReadLengthPrefixedUnicodeString(BinaryReader br)
        {
            OleObjectDataStreams.LengthPrefixedUnicodeString LPUniS = new LengthPrefixedUnicodeString();
            LPUniS.Length = br.ReadUInt32();
            LPUniS.String = BinaryHelper.GetString(br, LPUniS.Length, LPUniS.Encoding);
            return LPUniS;
        }

        private OleObjectDataStreams.LengthPrefixedAnsiString ReadLengthPrefixedAnsiString(BinaryReader br)
        {
            OleObjectDataStreams.LengthPrefixedAnsiString LPAnsiS = new LengthPrefixedAnsiString();
            LPAnsiS.Length = br.ReadUInt32();
            LPAnsiS.String = BinaryHelper.GetString(br, LPAnsiS.Length, LPAnsiS.Encoding);
            return LPAnsiS;
        }

        private OleObjectDataStreams.FILETIME ReadFILETIME(BinaryReader br)
        {
            OleObjectDataStreams.FILETIME FILETIME = new OleObjectDataStreams.FILETIME();
            FILETIME.dwLowDateTime = br.ReadUInt32();
            FILETIME.dwHighDateTime = br.ReadUInt32();
            return FILETIME;
        }

        private void ExportCompObjHeader(ExcelWorksheet ws, ref int ci, CompObjHeader header)
        {
            ws.Cells[2, ci++].Value = header.Reserved1;
            ws.Cells[2, ci++].Value = header.Version;
            ws.Cells[2, ci++].Value = header.Reserved2;
        }

        private void ExportCompObj(ExcelWorksheet ws, ref int ci)
        {
            ExportCompObjHeader(ws, ref ci, _oleDataStreams.CompObj.Header);
            ExportLengthPrefixedAnsiString(ws, ref ci, _oleDataStreams.CompObj.AnsiUserType);
            ExportClipboardFormatOrAnsiString(ws, ref ci, _oleDataStreams.CompObj.AnsiClipboardFormat);
            ExportLengthPrefixedAnsiString(ws, ref ci, _oleDataStreams.CompObj.Reserved1);
            ws.Cells[2, ci++].Value = _oleDataStreams.CompObj.UnicodeMarker;
            ExportLengthPrefixedUnicodeString(ws, ref ci, _oleDataStreams.CompObj.UnicodeUserType);
            ExportClipboardFormatOrUnicodeString(ws, ref ci, _oleDataStreams.CompObj.UnicodeClipboardFormat);
            ExportLengthPrefixedUnicodeString(ws, ref ci, _oleDataStreams.CompObj.Reserved2);
        }

        private OleObjectDataStreams.CompObjHeader ReadCompObjHeader(BinaryReader br)
        {
            OleObjectDataStreams.CompObjHeader header = new OleObjectDataStreams.CompObjHeader();
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
                _oleDataStreams.CompObj.Header = ReadCompObjHeader(br);

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.CompObj.AnsiUserType = ReadLengthPrefixedAnsiString(br);

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.CompObj.AnsiClipboardFormat = ReadClipboardFormatOrAnsiString(br);

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                LengthPrefixedAnsiString Reserved1 = ReadLengthPrefixedAnsiString(br);
                if (Reserved1.Length == 0 || Reserved1.Length > 0x00000028 || string.IsNullOrEmpty(Reserved1.String))
                {
                    //return;
                }

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                var UnicodeMarker = br.ReadUInt32();

                if (UnicodeMarker != 0x71B239F4)
                {
                    return;
                }
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.CompObj.UnicodeUserType = ReadLengthPrefixedUnicodeString(br);

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.CompObj.UnicodeClipboardFormat = ReadClipboardFormatOrUnicodeString(br);

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.CompObj.Reserved2 = ReadLengthPrefixedUnicodeString(br);
            }
        }

        private void ExportMonikerStream(ExcelWorksheet ws, ref int ci, MonikerStream MonikerStream)
        {
            ExportCLSID(ws, ref ci, MonikerStream.ClsId);
            ws.Cells[2, ci++].Value = MonikerStream.StreamData1;
            ws.Cells[2, ci++].Value = MonikerStream.StreamData2;
            ws.Cells[2, ci++].Value = MonikerStream.StreamData3;
            ws.Cells[2, ci++].Value = MonikerStream.StreamData4;
        }

        private void ExportCLSID(ExcelWorksheet ws, ref int ci, CLSID ClsId)
        {
            ws.Cells[2, ci++].Value = ClsId.Data1;
            ws.Cells[2, ci++].Value = ClsId.Data2;
            ws.Cells[2, ci++].Value = ClsId.Data3;
            ws.Cells[2, ci++].Value = ClsId.Data4;
        }

        private void ExportLengthPrefixedUnicodeString(ExcelWorksheet ws, ref int ci, OleObjectDataStreams.LengthPrefixedUnicodeString LPUniS)
        {
            ws.Cells[2, ci++].Value = LPUniS.Length;
            ws.Cells[2, ci++].Value = LPUniS.String;
        }

        private void ExportLengthPrefixedAnsiString(ExcelWorksheet ws, ref int ci, OleObjectDataStreams.LengthPrefixedAnsiString LPAnsiS)
        {
            ws.Cells[2, ci++].Value = LPAnsiS.Length;
            ws.Cells[2, ci++].Value = LPAnsiS.String;
        }

        private void ExportFILETIME(ExcelWorksheet ws, ref int ci, OleObjectDataStreams.FILETIME FILETIME)
        {
            ws.Cells[2, ci++].Value = FILETIME.dwLowDateTime;
            ws.Cells[2, ci++].Value = FILETIME.dwHighDateTime;
        }

        private void ExportOle(ExcelWorksheet ws, ref int ci)
        {
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.Version;
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.Flags;
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.LinkUpdateOption;
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.Reserved1;
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.ReservedMonikerStreamSize;
            if (_oleDataStreams.Ole.ReservedMonikerStreamSize != 0)
            {
                ExportMonikerStream(ws, ref ci, _oleDataStreams.Ole.ReservedMonikerStream);
            }
            else
            {
                ci += 5;
            }
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.RelativeSourceMonikerStreamSize;
            if (_oleDataStreams.Ole.RelativeSourceMonikerStreamSize != 0)
            {
                ExportMonikerStream(ws, ref ci, _oleDataStreams.Ole.RelativeSourceMonikerStream);
            }
            else
            {
                ci += 5;
            }
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.AbsoluteSourceMonikerStreamSize;
            if (_oleDataStreams.Ole.AbsoluteSourceMonikerStreamSize != 0)
            {
                ExportMonikerStream(ws, ref ci, _oleDataStreams.Ole.AbsoluteSourceMonikerStream);
            }
            else
            {
                ci += 5;
            }
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.ClsidIndicator;
            ExportCLSID(ws, ref ci, _oleDataStreams.Ole.Clsid);
            ExportLengthPrefixedUnicodeString(ws, ref ci, _oleDataStreams.Ole.ReservedDisplayName);
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.Reserved2;
            ExportFILETIME(ws, ref ci, _oleDataStreams.Ole.LocalUpdateTime);
            ExportFILETIME(ws, ref ci, _oleDataStreams.Ole.LocalCheckUpdateTime);
            ExportFILETIME(ws, ref ci, _oleDataStreams.Ole.RemoteUpdateTime);
        }

        private void ReadOleStream(byte[] oleBytes)
        {
            using (var ms = RecyclableMemory.GetStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);
                _oleDataStreams.Ole.Version = br.ReadUInt32();
                _oleDataStreams.Ole.Flags = br.ReadUInt32();
                _oleDataStreams.Ole.LinkUpdateOption = br.ReadUInt32();
                _oleDataStreams.Ole.Reserved1 = br.ReadUInt32();
                _oleDataStreams.Ole.ReservedMonikerStreamSize = br.ReadUInt32();
                if (_oleDataStreams.Ole.ReservedMonikerStreamSize != 0)
                    ReadMONIKERSTREAM(br, _oleDataStreams.Ole.ReservedMonikerStreamSize - 4);

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.Ole.RelativeSourceMonikerStreamSize = br.ReadUInt32();
                if (_oleDataStreams.Ole.RelativeSourceMonikerStreamSize != 0)
                    ReadMONIKERSTREAM(br, _oleDataStreams.Ole.RelativeSourceMonikerStreamSize - 4);

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.Ole.AbsoluteSourceMonikerStreamSize = br.ReadUInt32();
                if (_oleDataStreams.Ole.AbsoluteSourceMonikerStreamSize != 0)
                    ReadMONIKERSTREAM(br, _oleDataStreams.Ole.AbsoluteSourceMonikerStreamSize - 4);

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.Ole.ClsidIndicator = br.ReadUInt32();

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.Ole.Clsid = ReadCLSID(br);

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.Ole.ReservedDisplayName = ReadLengthPrefixedUnicodeString(br);

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.Ole.Reserved2 = br.ReadUInt32();

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.Ole.LocalUpdateTime = ReadFILETIME(br);
                _oleDataStreams.Ole.LocalCheckUpdateTime = ReadFILETIME(br);
                _oleDataStreams.Ole.RemoteUpdateTime = ReadFILETIME(br);
            }
        }

        private void ExportOleNative(ExcelWorksheet ws, ref int ci)
        {
            ws.Cells[2, ci++].Value = _oleDataStreams.OleNative.NativeDataSize;
            ws.Cells[2, ci++].Value = _oleDataStreams.OleNative.NativeData;
        }

        private void ReadOleNative(byte[] oleBytes)
        {
            using (var ms = RecyclableMemory.GetStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);
                _oleDataStreams.OleNative.NativeDataSize = br.ReadUInt32();
                _oleDataStreams.OleNative.NativeData = br.ReadBytes((int)_oleDataStreams.OleNative.NativeDataSize);
            }
        }

        //private void ReadOlePres(byte[] oleBytes, ExcelWorksheet ws, ref int ci)
        //{
        //    using (var ms = new MemoryStream(oleBytes))
        //    {
        //        BinaryReader br = new BinaryReader(ms);
        //        var AnsiClipboardFormatFormatOrAnsiString = ReadClipboardFormatOrAnsiString(br, ws, ref ci);
        //        if (br.BaseStream.Position >= br.BaseStream.Length)
        //            return;
        //        var TargetDeviceSize = br.ReadUInt32();
        //        ws.Cells[2, ci++].Value = TargetDeviceSize;
        //        if (TargetDeviceSize >= 0x00000004)
        //        {
        //            ReadDVTARGETDEVICE(br, TargetDeviceSize, ws, ref ci); //TargetDevice
        //        }
        //        var Aspect = br.ReadUInt32();
        //        var Lindex = br.ReadUInt32();
        //        var Advf = br.ReadUInt32();
        //        var Reserved1 = br.ReadUInt32();
        //        var Width = br.ReadUInt32();
        //        var Height = br.ReadUInt32();
        //        var Size = br.ReadUInt32();
        //        var Data = br.ReadBytes((int)Size);

        //        ws.Cells[2, ci++].Value = Aspect;
        //        ws.Cells[2, ci++].Value = Lindex;
        //        ws.Cells[2, ci++].Value = Advf;
        //        ws.Cells[2, ci++].Value = Reserved1;
        //        ws.Cells[2, ci++].Value = Width;
        //        ws.Cells[2, ci++].Value = Height;
        //        ws.Cells[2, ci++].Value = Size;
        //        ws.Cells[2, ci++].Value = Data;

        //        byte[] Reserved2 = new byte[] { };
        //        if (AnsiClipboardFormatFormatOrAnsiString.Length > 0 && BitConverter.ToUInt32(AnsiClipboardFormatFormatOrAnsiString, 0) == 0x00000003)
        //            Reserved2 = br.ReadBytes(18);

        //        ws.Cells[2, ci++].Value = Reserved2;

        //        var TocSignature = br.ReadUInt32();
        //        var TocCount = br.ReadUInt32();

        //        ws.Cells[2, ci++].Value = TocSignature;
        //        ws.Cells[2, ci++].Value = TocCount;

        //        if (TocSignature == 0x494E414 || TocCount == 0)
        //            return;

        //        int c2 = ci;
        //        for (int i = 0; i < TocCount; i++)
        //        {
        //            ReadTOCENTRY(br, ws, ref c2);
        //            ws.InsertRow(2, 1);
        //            c2 = ci;
        //            br.BaseStream.Position = br.BaseStream.Length;
        //            if (br.BaseStream.Position >= br.BaseStream.Length)
        //                return;
        //        }
        //    }
        //}

        internal void LoadDocument()
        {
            using var p = new ExcelPackage(@"C:\epplusTest\RESULTS.xlsx");

            var oleRel = _worksheet.Part.GetRelationship(_oleObject.RelationshipId);
            if (oleRel != null && oleRel.TargetUri.ToString().Contains(".bin"))
            {
                var oleObj = UriHelper.ResolvePartUri(oleRel.SourceUri, oleRel.TargetUri);
                var olePart = _worksheet._package.ZipPackage.GetPart(oleObj);
                var oleStream = (MemoryStream)olePart.GetStream(FileMode.Open, FileAccess.Read);
                _document = new CompoundDocument(oleStream);
                _oleDataStreams = new OleObjectDataStreams();
                if (_document.Storage.DataStreams.ContainsKey("\u0001Ole10Native"))
                {
                    _oleDataStreams.OleNative = new OleObjectDataStreams.OleNativeStream();
                    ReadOleNative(_document.Storage.DataStreams["\u0001Ole10Native"]);

                    var ws = p.Workbook.Worksheets["OleNative"];
                    ws.InsertRow(2, 1);
                    ws.Cells["A2"].Value = this._worksheet.Workbook._package.File.Name;
                    int colIndex = 2;
                    ExportOleNative(ws, ref colIndex);
                }
                if (_document.Storage.DataStreams.ContainsKey("\u0001Ole"))
                {
                    _oleDataStreams.Ole = new OleObjectDataStreams.OleObjectStream();
                    ReadOleStream(_document.Storage.DataStreams["\u0001Ole"]);

                    var ws = p.Workbook.Worksheets["Ole"];
                    ws.InsertRow(2, 1);
                    ws.Cells["A2"].Value = this._worksheet.Workbook._package.File.Name;
                    int colIndex = 2;
                    ExportOle(ws, ref colIndex);
                }
                if (_document.Storage.DataStreams.ContainsKey("\u0001CompObj"))
                {
                    _oleDataStreams.CompObj = new OleObjectDataStreams.CompObjStream();
                    ReadCompObjStream(_document.Storage.DataStreams["\u0001CompObj"]);

                    var ws = p.Workbook.Worksheets["CompObj"];
                    ws.InsertRow(2, 1);
                    ws.Cells["A2"].Value = this._worksheet.Workbook._package.File.Name;
                    int colIndex = 2;
                }
                //for (int i = 0; i <= 999; i++)
                //{
                //    string olePres = "\u0002OlePres" + i.ToString("D3");
                //    if (_document.Storage.DataStreams.ContainsKey(olePres))
                //    {
                //        var ws = p.Workbook.Worksheets["OlePres"];
                //        ws.InsertRow(2, 1);
                //        ws.Cells["A2"].Value = this._worksheet.Workbook._package.File.Name;
                //        int colIndex = 2;
                //        ReadOlePres(_document.Storage.DataStreams[olePres], ws, ref colIndex);
                //    }
                //}
            }
            p.Save();
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

    internal class OleObjectDataStreams
    {

        internal OleNativeStream OleNative;
        internal OleObjectStream Ole;
        internal CompObjStream CompObj;

        internal class MonikerStream
        {
            internal CLSID ClsId;
            internal UInt32 StreamData1;
            internal UInt16 StreamData2;
            internal UInt32 StreamData3; //Size of StreamData4
            internal string StreamData4;
            internal Encoding encoding = Encoding.Unicode;
        }

        internal class CLSID
        {
            internal UInt32 Data1;
            internal UInt16 Data2;
            internal UInt16 Data3;
            internal UInt64 Data4;
        }

        internal class LengthPrefixedUnicodeString
        {
            internal UInt32 Length;
            internal string String;
            internal Encoding Encoding = Encoding.Unicode;
        }

        internal class LengthPrefixedAnsiString
        {
            internal UInt32 Length;
            internal string String;
            internal Encoding Encoding = Encoding.ASCII;
        }

        internal class ClipboardFormatOrUnicodeString
        {
            //If this is set to 0x00000000, the FormatOrUnicodeString field MUST
            //NOT be present.If this is set to 0xffffffff or 0xfffffffe, the FormatOrUnicodeString field MUST be
            //4 bytes in size and MUST contain a standard clipboard format identifier
            //Otherwise, the FormatOrUnicodeString field MUST be set to a Unicode string containing the name of a registered clipboard format
            //and the MarkerOrLength field MUST be set to the number of Unicode characters in the FormatOrUnicodeString field, including the
            //terminating null character.
            internal UInt32 MarkerOrLength;
            internal Byte[] FormatOrUnicodeString;
        }

        internal class ClipboardFormatOrAnsiString
        {
            //If this field is set to 0xFFFFFFFF or 0xFFFFFFFE,
            //the FormatOrAnsiString field MUST be 4 bytes in size and MUST contain a standard clipboard format identifier.
            //If this set to a value other than 0x00000000,
            //the FormatOrAnsiString field MUST be set to a null-terminated ANSI string containing the name of a registered clipboard format
            internal UInt32 MarkerOrLength;
            internal Byte[] FormatOrAnsiString;
        }

        internal class FILETIME
        {
            internal UInt32 dwLowDateTime;
            internal UInt32 dwHighDateTime;
        }

        internal class OleObjectStream
        {
            internal UInt32 Version;
            internal UInt32 Flags;
            internal UInt32 LinkUpdateOption;
            internal UInt32 Reserved1;
            internal UInt32 ReservedMonikerStreamSize; //Subtract by 4 when reading if not 0
            internal MonikerStream ReservedMonikerStream;

            internal UInt32 RelativeSourceMonikerStreamSize; //Subtract by 4 when reading if not 0
            internal MonikerStream RelativeSourceMonikerStream;

            internal UInt32 AbsoluteSourceMonikerStreamSize; //Subtract by 4 when reading if not 0
            internal MonikerStream AbsoluteSourceMonikerStream;

            internal UInt32 ClsidIndicator;
            internal CLSID Clsid;

            internal LengthPrefixedUnicodeString ReservedDisplayName;

            internal UInt32 Reserved2;

            internal FILETIME LocalUpdateTime;
            internal FILETIME LocalCheckUpdateTime;
            internal FILETIME RemoteUpdateTime;
        }

        internal class CompObjHeader
        {
            internal UInt32 Reserved1;
            internal UInt32 Version;
            internal byte[] Reserved2 = new byte[20];
        }

        internal class CompObjStream
        {
            internal CompObjHeader Header;
            internal LengthPrefixedAnsiString AnsiUserType;
            internal ClipboardFormatOrAnsiString AnsiClipboardFormat; //MarkerOrLength field of the ClipboardFormatOrAnsiString structure contains a value other than 0x00000000, 0xffffffff, or 0xfffffffe, the value MUST NOT be greater than 0x00000190. Otherwise the CompObjStream structure is invalid.
            internal LengthPrefixedAnsiString Reserved1;
            //      Reserved1 (variable): If present, this MUST be a LengthPrefixedAnsiString structure.
            //      If the Length field of the LengthPrefixedAnsiString contains a value of 0 or a value that is greater than 0x00000028,
            //      the remaining fields of the structure starting with the String field of the LengthPrefixedAnsiString MUST be ignored on processing.
            //      If the String field of the LengthPrefixedAnsiString is not present, the remaining fields of the
            //      structure starting with the UnicodeMarker field MUST be ignored on processing.
            //      Otherwise, the String field of the LengthPrefixedAnsiString MUST be ignored on processing.
            internal UInt32 UnicodeMarker; //If this field is present and is NOT set to 0x71B239F4, the remaining fields of the structure MUST be ignored on processing.
            internal LengthPrefixedUnicodeString UnicodeUserType;
            internal ClipboardFormatOrUnicodeString UnicodeClipboardFormat; //MarkerOrLength field of the ClipboardFormatOrUnicodeString structure contains a value other than 0x00000000, 0xffffffff, or 0xfffffffe, the value MUST NOT be more than 0x00000190. Otherwise, the CompObjStream structure is invalid.
            internal LengthPrefixedUnicodeString Reserved2;
        }

        internal class OleNativeStream
        {
            internal UInt32 NativeDataSize;
            internal byte[] NativeData;
        }
    }

}

/*
 * TODO:
 * Skapa struktur med klasser för Ole, CompObj och OleNative
 * Läs till dessa
 * Skapa default värden för aString och Resereved1String i CompObj
 * Funktion för att sätta StreamData4 i Ole som är worksheetName!ObjectName
 * 
 * 
 * Skapa egen EmbeddedObject
 * user specidified aString och Reserved1String
 * user specified Image
 * user specified file
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
 */