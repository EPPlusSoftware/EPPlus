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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;


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

        private byte[] ReadClipboardFormatOrAnsiString(BinaryReader br, ExcelWorksheet ws, ref int ci)
        {
            var MarkerOrLength = br.ReadUInt32();
            ws.Cells[2, ci++].Value = MarkerOrLength;
            byte[] FormatOrAnsiString = null;
            if (MarkerOrLength > 0x00000190 || MarkerOrLength == 0x00000000)
            {
                ws.Cells[2, ci++].Value = "";
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
            ws.Cells[2, ci++].Value = FormatOrAnsiString;
            return FormatOrAnsiString;
        }

        private byte[] ReadClipboardFormatOrUnicodeString(BinaryReader br, ExcelWorksheet ws, ref int ci)
        {
            var MarkerOrLength = br.ReadUInt32();
            ws.Cells[2, ci++].Value = MarkerOrLength;
            byte[] FormatOrUnicodeString = null;
            if (MarkerOrLength > 0x00000190 || MarkerOrLength == 0x00000000)
            {
                ws.Cells[2, ci++].Value = "";
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
            ws.Cells[2, ci++].Value = FormatOrUnicodeString;
            return FormatOrUnicodeString;
        }

        private void ReadTOCENTRY(BinaryReader br, ExcelWorksheet ws, ref int ci)
        {
            ReadClipboardFormatOrAnsiString(br, ws, ref ci); //AnsiClipboardFormat
            var TargetDeviceSize = br.ReadUInt32();
            var Aspect = br.ReadUInt32();
            var Lindex = br.ReadUInt32();
            var Tymed = br.ReadUInt32();
            var Reserved1 = br.ReadBytes(12);
            var Advf = br.ReadUInt32();
            var Reserved2 = br.ReadUInt32();

            ws.Cells[2, ci++].Value = TargetDeviceSize;
            ws.Cells[2, ci++].Value = Aspect;
            ws.Cells[2, ci++].Value = Lindex;
            ws.Cells[2, ci++].Value = Tymed;
            ws.Cells[2, ci++].Value = Reserved1;
            ws.Cells[2, ci++].Value = Advf;
            ws.Cells[2, ci++].Value = Reserved2;

            ReadDVTARGETDEVICE(br, TargetDeviceSize, ws, ref ci); //TargetDevice
        }

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
                ci += 33;
        }
        private void ReadMONIKERSTREAM(BinaryReader br, uint size, ExcelWorksheet ws, ref int ci)
        {
            var ClsId = br.ReadBytes(16);
            var StreamData1 = br.ReadUInt32();
            var StreamData2 = br.ReadUInt16();
            var StreamData3 = br.ReadUInt32();
            var StreamData4 = BinaryHelper.GetString(br, StreamData3, Encoding.ASCII);

            ws.Cells[2, ci++].Value = ClsId;
            ws.Cells[2, ci++].Value = StreamData1;
            ws.Cells[2, ci++].Value = StreamData2;
            ws.Cells[2, ci++].Value = StreamData3;
            ws.Cells[2, ci++].Value = StreamData4;

        }

        private void ReadCLSID(BinaryReader br, ExcelWorksheet ws, ref int ci)
        {
            var Data1 = br.ReadUInt32();
            var Data2 = br.ReadUInt16();
            var Data3 = br.ReadUInt16();
            var Data4 = br.ReadUInt64();

            ws.Cells[2, ci++].Value = Data1;
            ws.Cells[2, ci++].Value = Data2;
            ws.Cells[2, ci++].Value = Data3;
            ws.Cells[2, ci++].Value = Data4;
        }

        private void ReadLengthPrefixedUnicodeString(BinaryReader br, ExcelWorksheet ws, ref int ci)
        {
            var Length = br.ReadUInt32();
            var uString = BinaryHelper.GetString(br, Length, Encoding.Unicode);

            ws.Cells[2, ci++].Value = Length;
            ws.Cells[2, ci++].Value = uString;
        }

        private void ReadLengthPrefixedAnsiString(BinaryReader br, ExcelWorksheet ws, ref int ci)
        {
            var Length = br.ReadUInt32();
            var aString = BinaryHelper.GetString(br, Length, Encoding.ASCII);

            ws.Cells[2, ci++].Value = Length;
            ws.Cells[2, ci++].Value = aString;
        }

        private void ReadFILETIME(BinaryReader br, ExcelWorksheet ws, ref int ci)
        {
            var dwLowDateTime = br.ReadUInt32();
            var dwHighDateTime = br.ReadUInt32();

            ws.Cells[2, ci++].Value = dwLowDateTime;
            ws.Cells[2, ci++].Value = dwHighDateTime;
        }

        private void ReadCompObjHeader(BinaryReader br, ExcelWorksheet ws, ref int ci)
        {
            var Reserved1 = br.ReadUInt32();
            var Version = br.ReadUInt32();
            var Reserved2 = br.ReadBytes(20);

            ws.Cells[2, ci++].Value = Reserved1;
            ws.Cells[2, ci++].Value = Version;
            ws.Cells[2, ci++].Value = Reserved2;
        }

        private void ReadCompObjStream(byte[] oleBytes, ExcelWorksheet ws, ref int ci)
        {
            using (var ms = RecyclableMemory.GetStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);

                ReadCompObjHeader(br, ws, ref ci); //Header

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadLengthPrefixedAnsiString(br, ws, ref ci); //AnsiUserType

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadClipboardFormatOrAnsiString(br, ws, ref ci); //AnsiClipboardFormat 

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                //Reserved1 should be a LengthPrefixedUnicodeString
                var Reserved1Length = br.ReadUInt32();
                string Reserved1String = "";
                if (Reserved1Length == 0 || Reserved1Length > 0x00000028)
                {
                    //return;
                }
                else
                {
                    Reserved1String = BinaryHelper.GetString(br, Reserved1Length, Encoding.ASCII);
                }

                ws.Cells[2, ci++].Value = Reserved1Length;
                ws.Cells[2, ci++].Value = Reserved1String;

                if (string.IsNullOrEmpty(Reserved1String))
                {
                    //return;
                }
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                var UnicodeMarker = br.ReadUInt32();
                ws.Cells[2, ci++].Value = UnicodeMarker;
                if (UnicodeMarker != 0x71B239F4)
                {
                    return;
                }
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadLengthPrefixedUnicodeString(br, ws, ref ci); //UnicodeUserType
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadClipboardFormatOrUnicodeString(br, ws, ref ci);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadLengthPrefixedUnicodeString(br, ws, ref ci); //Reserved2
            }
        }

        private void ReadOleStream(byte[] oleBytes, ExcelWorksheet ws, ref int ci)
        {
            using (var ms = RecyclableMemory.GetStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);
                var Version = br.ReadUInt32();
                var Flags = br.ReadUInt32();
                var LinkUpdateOption = br.ReadUInt32();
                var Reserved1 = br.ReadUInt32();
                var ReservedMonikerStreamSize = br.ReadUInt32() - 4;
                ws.Cells[2, ci++].Value = Version;
                ws.Cells[2, ci++].Value = Flags;
                ws.Cells[2, ci++].Value = LinkUpdateOption;
                ws.Cells[2, ci++].Value = Reserved1;
                ws.Cells[2, ci++].Value = ReservedMonikerStreamSize;
                ws.Cells[2, ci].Value = "";

                if (ReservedMonikerStreamSize != 0)
                {
                    ReadMONIKERSTREAM(br, ReservedMonikerStreamSize, ws, ref ci);
                }
                else
                {
                    ci += 5;
                }

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                var RelativeSourceMonikerStreamSize = br.ReadUInt32() - 4;
                ws.Cells[2, ci++].Value = RelativeSourceMonikerStreamSize;
                ws.Cells[2, ci].Value = "";
                if (RelativeSourceMonikerStreamSize != 0)
                {
                    ReadMONIKERSTREAM(br, RelativeSourceMonikerStreamSize, ws, ref ci);
                }
                else
                {
                    ci += 5;
                }

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                var AbsoluteSourceMonikerStreamSize = br.ReadUInt32() - 4;
                ws.Cells[2, ci++].Value = AbsoluteSourceMonikerStreamSize;
                ws.Cells[2, ci].Value = "";
                if (AbsoluteSourceMonikerStreamSize != 0)
                {
                    ReadMONIKERSTREAM(br, AbsoluteSourceMonikerStreamSize, ws, ref ci);
                }
                else
                {
                    ci += 5;
                }

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                var ClsidIndicator = br.ReadUInt32();
                ws.Cells[2, ci++].Value = ClsidIndicator;

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadCLSID(br, ws, ref ci); //Clsid

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadLengthPrefixedUnicodeString(br, ws, ref ci); //ReservedDisplayName

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                var Reserved2 = br.ReadUInt32();
                ws.Cells[2, ci++].Value = Reserved2;

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                ReadFILETIME(br, ws, ref ci); //LocalUpdateTime
                ReadFILETIME(br, ws, ref ci); //LocalCheckUpdateTime
                ReadFILETIME(br, ws, ref ci); //RemoteUpdateTime
            }
        }

        private void ReadOleNative(byte[] oleBytes, ExcelWorksheet ws, ref int ci)
        {
            using (var ms = RecyclableMemory.GetStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);
                var NativeDataSize = br.ReadUInt32();
                var NativeData = br.ReadBytes((int)NativeDataSize);

                ws.Cells[2, ci++].Value = NativeDataSize;
                ws.Cells[2, ci++].Value = NativeData;
            }
        }

        private void ReadOlePres(byte[] oleBytes, ExcelWorksheet ws, ref int ci)
        {
            using (var ms = new MemoryStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);
                var AnsiClipboardFormatFormatOrAnsiString = ReadClipboardFormatOrAnsiString(br, ws, ref ci);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;
                var TargetDeviceSize = br.ReadUInt32();
                ws.Cells[2, ci++].Value = TargetDeviceSize;
                if (TargetDeviceSize >= 0x00000004)
                {
                    ReadDVTARGETDEVICE(br, TargetDeviceSize, ws, ref ci); //TargetDevice
                }
                var Aspect = br.ReadUInt32();
                var Lindex = br.ReadUInt32();
                var Advf = br.ReadUInt32();
                var Reserved1 = br.ReadUInt32();
                var Width = br.ReadUInt32();
                var Height = br.ReadUInt32();
                var Size = br.ReadUInt32();
                var Data = br.ReadBytes((int)Size);

                ws.Cells[2, ci++].Value = Aspect;
                ws.Cells[2, ci++].Value = Lindex;
                ws.Cells[2, ci++].Value = Advf;
                ws.Cells[2, ci++].Value = Reserved1;
                ws.Cells[2, ci++].Value = Width;
                ws.Cells[2, ci++].Value = Height;
                ws.Cells[2, ci++].Value = Size;
                ws.Cells[2, ci++].Value = Data;

                byte[] Reserved2 = new byte[] { };
                if ( AnsiClipboardFormatFormatOrAnsiString.Length > 0 && BitConverter.ToUInt32(AnsiClipboardFormatFormatOrAnsiString, 0) == 0x00000003)
                    Reserved2 = br.ReadBytes(18);

                ws.Cells[2, ci++].Value = Reserved2;

                var TocSignature = br.ReadUInt32();
                var TocCount = br.ReadUInt32();

                ws.Cells[2, ci++].Value = TocSignature;
                ws.Cells[2, ci++].Value = TocCount;

                if (TocSignature == 0x494E414 || TocCount == 0)
                    return;
                for (int i = 0; i < TocCount; i++)
                {
                    ReadTOCENTRY(br, ws, ref ci);
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
            using var p = new ExcelPackage(@"C:\epplusTest\RESULTS.xlsx");

            var oleRel = _worksheet.Part.GetRelationship(_oleObject.RelationshipId);
            if (oleRel != null && oleRel.TargetUri.ToString().Contains(".bin"))
            {
                var oleObj = UriHelper.ResolvePartUri(oleRel.SourceUri, oleRel.TargetUri);
                var olePart = _worksheet._package.ZipPackage.GetPart(oleObj);
                var oleStream = (MemoryStream)olePart.GetStream(FileMode.Open, FileAccess.Read);
                _document = new CompoundDocument(oleStream);
                if (_document.Storage.DataStreams.ContainsKey("\u0001Ole10Native"))
                {
                    var ws = p.Workbook.Worksheets["OleNative"];
                    ws.InsertRow(2, 1);
                    ws.Cells["A2"].Value = this._worksheet.Workbook._package.File.Name;
                    int colIndex = 2;
                    ReadOleNative(_document.Storage.DataStreams["\u0001Ole10Native"], ws, ref colIndex);
                }
                if (_document.Storage.DataStreams.ContainsKey("\u0001Ole"))
                {
                    var ws = p.Workbook.Worksheets["Ole"];
                    ws.InsertRow(2, 1);
                    ws.Cells["A2"].Value = this._worksheet.Workbook._package.File.Name;
                    int colIndex = 2;
                    ReadOleStream(_document.Storage.DataStreams["\u0001Ole"], ws, ref colIndex);
                }
                if (_document.Storage.DataStreams.ContainsKey("\u0001CompObj"))
                {
                    var ws = p.Workbook.Worksheets["CompObj"];
                    ws.InsertRow(2, 1);
                    ws.Cells["A2"].Value = this._worksheet.Workbook._package.File.Name;
                    int colIndex = 2;
                    ReadCompObjStream(_document.Storage.DataStreams["\u0001CompObj"], ws, ref colIndex);
                }
                for (int i = 0; i <= 999; i++)
                {
                    string olePres = "\u0002OlePres" + i.ToString("D3");
                    if (_document.Storage.DataStreams.ContainsKey(olePres))
                    {
                        var ws = p.Workbook.Worksheets["OlePres"];
                        ws.InsertRow(2, 1);
                        ws.Cells["A2"].Value = this._worksheet.Workbook._package.File.Name;
                        int colIndex = 2;
                        ReadOlePres(_document.Storage.DataStreams[olePres], ws, ref colIndex);
                    }
                }
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