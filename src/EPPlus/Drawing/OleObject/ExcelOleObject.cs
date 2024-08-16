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
using static OfficeOpenXml.Drawing.OleObject.OleObjectDataStreams;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;


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
                LoadEmbeddedDocument();
            }
            else
            {
                isExternalLink = true;
                LoadExternalLink();
            }
        }

        internal ExcelOleObject(ExcelDrawings drawings, XmlNode node, string filePath, bool link, string mediaFilePath = "", ExcelGroupShape parent = null)
            : base(drawings, node, "xdr:sp", "xdr:nvSpPr/xdr:cNvPr", parent)
        {
            _worksheet = drawings.Worksheet;
            string relId = "";
            if (link)
            {
                isExternalLink = true;
                //create ExternalLink
            }
            else
            {
                isExternalLink = false;
                relId = EmbedDocument(filePath);
            }

            //Create Media
            //User supplied picture or our own placeholder
            //Construct icon with rectable with txbody set to filename and an autorectangle. Somehow you can't see the txbody or autorectangle when icon is complete. only when you ungroup.
            //create Uri
            //Create relationship
            //read bytes from filepath
            //same as bin files?
            int newID = 1;
            var Uri = GetNewUri(_worksheet._package.ZipPackage, "/xl/media/image{0}.emf", ref newID);
            var part = _worksheet._package.ZipPackage.CreatePart(Uri, "image/x-emf", CompressionLevel.None, "emf");
            var rel = _worksheet.Part.CreateRelationship(Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
            byte[] image = File.ReadAllBytes(mediaFilePath);
            MemoryStream ms = (MemoryStream)part.GetStream(FileMode.Create, FileAccess.Write);
            ms.Write(image, 0, image.Length);
            var imgRelId = rel.Id;

            //Create drawings xml
            string name = _drawings.GetUniqueDrawingName("Object 1");
            XmlElement spElement = CreateShapeNode();
            spElement.InnerXml = CreateOleObjectDrawingNode(name);
            CreateClientData();
            From.Column = 0; From.ColumnOff = 0;
            From.Row = 0; From.RowOff = 0;
            To.Column = 1; To.ColumnOff = 171450;
            To.Row = 2; To.RowOff = 133350;

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
            sb.AppendFormat("<oleObject progId=\"{0}\" shapeId=\"{1}\" r:id=\"{2}\">", "Packager Shell Object"/*_oleDataStreams.CompObj.Reserved1.String*/,  _id, relId);
            sb.AppendFormat("<objectPr defaultSize=\"0\" r:id=\"{0}\">", imgRelId); //SET relId TO MEDIA HERE autoPict=\"0\"
            sb.Append("<anchor moveWithCells=\"1\">");
            sb.Append("<from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></from>");       //SET VALUE BASED ON MEDIA
            sb.Append("<to><xdr:col>1</xdr:col><xdr:colOff>171450</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>133350</xdr:rowOff></to>"); //SET VALUE BASED ON MEDIA
            sb.Append("</anchor></objectPr></oleObject>");
            sb.Append("</mc:Choice>");
            //fallback
            sb.AppendFormat("<mc:Fallback><oleObject progId=\"{0}\" shapeId=\"{1}\" r:id=\"{2}\" />", "Packager Shell Object" /*_oleDataStreams.CompObj.Reserved1.String*/, _id, relId);
            sb.Append("</mc:Fallback></mc:AlternateContent>");
            wsNode.InnerXml = sb.ToString();
            var oleObjectNode = wsNode.GetChildAtPosition(0).GetChildAtPosition(0);
            _oleObject = new OleObjectInternal(_worksheet.NameSpaceManager, oleObjectNode);
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

        #region Export

        private void ExportLengthPrefixedUnicodeString(ExcelWorksheet ws, ref int ci, OleObjectDataStreams.LengthPrefixedUnicodeString LPUniS)
        {
            if (LPUniS == null)
            {
                ci += 2;
                return;
            }
            ws.Cells[2, ci++].Value = LPUniS.Length;
            ws.Cells[2, ci++].Value = LPUniS.String;
        }

        private void ExportLengthPrefixedAnsiString(ExcelWorksheet ws, ref int ci, OleObjectDataStreams.LengthPrefixedAnsiString LPAnsiS)
        {
            if (LPAnsiS == null)
            {
                ci += 2;
                return;
            }
            ws.Cells[2, ci++].Value = LPAnsiS.Length;
            ws.Cells[2, ci++].Value = LPAnsiS.String;
        }

        private void ExportClipboardFormatOrUnicodeString(ExcelWorksheet ws, ref int ci, OleObjectDataStreams.ClipboardFormatOrUnicodeString CFOUS)
        {
            if (CFOUS == null)
            {
                ci += 2;
                return;
            }
            ws.Cells[2, ci++].Value = CFOUS.MarkerOrLength;
            ws.Cells[2, ci++].Value = CFOUS.FormatOrUnicodeString;
        }

        private void ExportClipboardFormatOrAnsiString(ExcelWorksheet ws, ref int ci, OleObjectDataStreams.ClipboardFormatOrAnsiString CFOAS)
        {
            if (CFOAS == null)
            {
                ci += 2;
                return;
            }
            ws.Cells[2, ci++].Value = CFOAS.MarkerOrLength;
            ws.Cells[2, ci++].Value = CFOAS.FormatOrAnsiString;
        }

        private void ExportCLSID(ExcelWorksheet ws, ref int ci, CLSID ClsId)
        {
            if (ClsId == null)
            {
                ci += 4;
                return;
            }
            ws.Cells[2, ci++].Value = ClsId.Data1;
            ws.Cells[2, ci++].Value = ClsId.Data2;
            ws.Cells[2, ci++].Value = ClsId.Data3;
            ws.Cells[2, ci++].Value = ClsId.Data4;
        }

        private void ExportMonikerStream(ExcelWorksheet ws, ref int ci, MonikerStream MonikerStream)
        {
            if (MonikerStream == null)
            {
                ci += 8;
                return;
            }
            ExportCLSID(ws, ref ci, MonikerStream.ClsId);
            ws.Cells[2, ci++].Value = MonikerStream.StreamData1;
            ws.Cells[2, ci++].Value = MonikerStream.StreamData2;
            ws.Cells[2, ci++].Value = MonikerStream.StreamData3;
            ws.Cells[2, ci++].Value = MonikerStream.StreamData4;
        }

        private void ExportFILETIME(ExcelWorksheet ws, ref int ci, OleObjectDataStreams.FILETIME FILETIME)
        {
            if (FILETIME == null)
            {
                ci += 2;
                return;
            }
            ws.Cells[2, ci++].Value = FILETIME.dwLowDateTime;
            ws.Cells[2, ci++].Value = FILETIME.dwHighDateTime;
        }

        private void ExportCompObjHeader(ExcelWorksheet ws, ref int ci, CompObjHeader header)
        {
            if (header == null)
            {
                ci += 3;
                return;
            }
            ws.Cells[2, ci++].Value = header.Reserved1;
            ws.Cells[2, ci++].Value = header.Version;
            ws.Cells[2, ci++].Value = header.Reserved2;
        }

        private void ExportCompObj(ExcelWorksheet ws, ref int ci)
        {
            if (_oleDataStreams.CompObj == null)
                return;
            ExportCompObjHeader(ws, ref ci, _oleDataStreams.CompObj.Header);
            ExportLengthPrefixedAnsiString(ws, ref ci, _oleDataStreams.CompObj.AnsiUserType);
            ExportClipboardFormatOrAnsiString(ws, ref ci, _oleDataStreams.CompObj.AnsiClipboardFormat);
            ExportLengthPrefixedAnsiString(ws, ref ci, _oleDataStreams.CompObj.Reserved1);
            ws.Cells[2, ci++].Value = _oleDataStreams.CompObj.UnicodeMarker;
            ExportLengthPrefixedUnicodeString(ws, ref ci, _oleDataStreams.CompObj.UnicodeUserType);
            ExportClipboardFormatOrUnicodeString(ws, ref ci, _oleDataStreams.CompObj.UnicodeClipboardFormat);
            ExportLengthPrefixedUnicodeString(ws, ref ci, _oleDataStreams.CompObj.Reserved2);
        }

        private void ExportOle(ExcelWorksheet ws, ref int ci)
        {
            if (_oleDataStreams.Ole == null)
                return;
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.Version;
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.Flags;
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.LinkUpdateOption;
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.Reserved1;
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.ReservedMonikerStreamSize;
            ExportMonikerStream(ws, ref ci, _oleDataStreams.Ole.ReservedMonikerStream);
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.RelativeSourceMonikerStreamSize;
            ExportMonikerStream(ws, ref ci, _oleDataStreams.Ole.RelativeSourceMonikerStream);
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.AbsoluteSourceMonikerStreamSize;
            ExportMonikerStream(ws, ref ci, _oleDataStreams.Ole.AbsoluteSourceMonikerStream);
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.ClsIdIndicator;
            ExportCLSID(ws, ref ci, _oleDataStreams.Ole.ClsId);
            ExportLengthPrefixedUnicodeString(ws, ref ci, _oleDataStreams.Ole.ReservedDisplayName);
            ws.Cells[2, ci++].Value = _oleDataStreams.Ole.Reserved2;
            ExportFILETIME(ws, ref ci, _oleDataStreams.Ole.LocalUpdateTime);
            ExportFILETIME(ws, ref ci, _oleDataStreams.Ole.LocalCheckUpdateTime);
            ExportFILETIME(ws, ref ci, _oleDataStreams.Ole.RemoteUpdateTime);
        }

        private void ExportOleNative(ExcelWorksheet ws, ref int ci)
        {
            if (_oleDataStreams.OleNative == null)
                return;
            ws.Cells[2, ci++].Value = _oleDataStreams.OleNative.NativeDataSize;
            ws.Cells[2, ci++].Value = _oleDataStreams.OleNative.NativeData;
        }
        #endregion

        #region ReadBinaries

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
            LPAnsiS.String = BinaryHelper.GetString(br, LPAnsiS.Length, LPAnsiS.Encoding).Trim('\0');
            return LPAnsiS;
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

        private OleObjectDataStreams.ClipboardFormatOrAnsiString ReadClipboardFormatOrAnsiString(BinaryReader br)
        {
            OleObjectDataStreams.ClipboardFormatOrAnsiString CFOAS = new OleObjectDataStreams.ClipboardFormatOrAnsiString();
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

        private OleObjectDataStreams.CLSID ReadCLSID(BinaryReader br)
        {
            OleObjectDataStreams.CLSID CLSID = new OleObjectDataStreams.CLSID();
            CLSID.Data1 = br.ReadUInt32();
            CLSID.Data2 = br.ReadUInt16();
            CLSID.Data3 = br.ReadUInt16();
            CLSID.Data4 = br.ReadUInt64();
            return CLSID;
        }

        private OleObjectDataStreams.MonikerStream ReadMONIKERSTREAM(BinaryReader br, uint size)
        {
            OleObjectDataStreams.MonikerStream monikerStream = new OleObjectDataStreams.MonikerStream();
            monikerStream.ClsId = ReadCLSID(br);
            monikerStream.StreamData1 = br.ReadUInt32();
            monikerStream.StreamData2 = br.ReadUInt16();
            monikerStream.StreamData3 = br.ReadUInt32();
            monikerStream.StreamData4 = BinaryHelper.GetString(br, monikerStream.StreamData3, Encoding.ASCII);
            return monikerStream;
        }

        private OleObjectDataStreams.FILETIME ReadFILETIME(BinaryReader br)
        {
            OleObjectDataStreams.FILETIME FILETIME = new OleObjectDataStreams.FILETIME();
            FILETIME.dwLowDateTime = br.ReadUInt32();
            FILETIME.dwHighDateTime = br.ReadUInt32();
            return FILETIME;
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

                _oleDataStreams.CompObj.Reserved1 = ReadLengthPrefixedAnsiString(br);
                if (_oleDataStreams.CompObj.Reserved1.Length == 0 || _oleDataStreams.CompObj.Reserved1.Length > 0x00000028 || string.IsNullOrEmpty(_oleDataStreams.CompObj.Reserved1.String))
                {
                    //throw error, invaldi comp obj file
                }

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.CompObj.UnicodeMarker = br.ReadUInt32();

                //if (_oleDataStreams.CompObj.UnicodeMarker != 0x71B239F4)
                //{
                //    return;
                //}
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
                    _oleDataStreams.Ole.ReservedMonikerStream = ReadMONIKERSTREAM(br, _oleDataStreams.Ole.ReservedMonikerStreamSize - 4);

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.Ole.RelativeSourceMonikerStreamSize = br.ReadUInt32();
                if (_oleDataStreams.Ole.RelativeSourceMonikerStreamSize != 0)
                    _oleDataStreams.Ole.RelativeSourceMonikerStream = ReadMONIKERSTREAM(br, _oleDataStreams.Ole.RelativeSourceMonikerStreamSize - 4);

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.Ole.AbsoluteSourceMonikerStreamSize = br.ReadUInt32();
                if (_oleDataStreams.Ole.AbsoluteSourceMonikerStreamSize != 0)
                    _oleDataStreams.Ole.AbsoluteSourceMonikerStream = ReadMONIKERSTREAM(br, _oleDataStreams.Ole.AbsoluteSourceMonikerStreamSize - 4);

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.Ole.ClsIdIndicator = br.ReadUInt32();

                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStreams.Ole.ClsId = ReadCLSID(br);

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

        private void ReadOleNative(byte[] oleBytes)
        {
            using (var ms = RecyclableMemory.GetStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);
                _oleDataStreams.OleNative.NativeDataSize = br.ReadUInt32();
                _oleDataStreams.OleNative.NativeData = br.ReadBytes((int)_oleDataStreams.OleNative.NativeDataSize);
            }
        }
        #endregion

        internal void LoadEmbeddedDocument()
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
                    ExportCompObj(ws, ref colIndex);
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

        #region WriteBinaries

        private byte[] ConcatenateByteArrays(params byte[][] arrays)
        {
            int dataLength = 0;
            foreach(var arr in arrays)
            {
                dataLength += arr.Length;
            }
            byte[] dataArray = new byte[dataLength];
            int offset = 0;
            foreach (var arr in arrays)
            {
                Buffer.BlockCopy(arr, 0, dataArray, offset, arr.Length);
                offset += arr.Length;
            }
            return dataArray;
        }

        private void CreateOleObject()
        {
            _oleDataStreams.Ole = new OleObjectStream();
            _oleDataStreams.Ole.ReservedMonikerStream = new MonikerStream();
            _oleDataStreams.Ole.ReservedMonikerStream.ClsId = new CLSID();
            byte[] size = ConcatenateByteArrays(BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStream.ClsId.Data1),
                                                BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStream.ClsId.Data2),
                                                BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStream.ClsId.Data3),
                                                BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStream.ClsId.Data4),
                                                BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStream.StreamData1),
                                                BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStream.StreamData2),
                                                BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStream.StreamData3),
                                                BinaryHelper.GetByteArray(_oleDataStreams.Ole.ReservedMonikerStream.StreamData4, _oleDataStreams.Ole.ReservedMonikerStream.Encoding));
            _oleDataStreams.Ole.ReservedMonikerStreamSize = (UInt32)size.Length;
            _oleDataStreams.Ole.RelativeSourceMonikerStream = new MonikerStream();
            _oleDataStreams.Ole.RelativeSourceMonikerStream.ClsId = new CLSID();
            size = ConcatenateByteArrays(BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStream.ClsId.Data1),
                                         BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStream.ClsId.Data2),
                                         BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStream.ClsId.Data3),
                                         BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStream.ClsId.Data4),
                                         BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStream.StreamData1),
                                         BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStream.StreamData2),
                                         BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStream.StreamData3),
                                         BinaryHelper.GetByteArray(_oleDataStreams.Ole.RelativeSourceMonikerStream.StreamData4, _oleDataStreams.Ole.RelativeSourceMonikerStream.Encoding));
            _oleDataStreams.Ole.RelativeSourceMonikerStreamSize = (UInt32)size.Length;
            _oleDataStreams.Ole.AbsoluteSourceMonikerStream = new MonikerStream();
            _oleDataStreams.Ole.AbsoluteSourceMonikerStream.ClsId = new CLSID();
            size = ConcatenateByteArrays(BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.ClsId.Data1),
                                         BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.ClsId.Data2),
                                         BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.ClsId.Data3),
                                         BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.ClsId.Data4),
                                         BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.StreamData1),
                                         BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.StreamData2),
                                         BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.StreamData3),
                                         BinaryHelper.GetByteArray(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.StreamData4, _oleDataStreams.Ole.AbsoluteSourceMonikerStream.Encoding));
            _oleDataStreams.Ole.AbsoluteSourceMonikerStreamSize = (UInt32)size.Length;
            _oleDataStreams.Ole.ClsId = new CLSID();
            _oleDataStreams.Ole.ReservedDisplayName = new LengthPrefixedUnicodeString();
            _oleDataStreams.Ole.LocalUpdateTime = new FILETIME();
            _oleDataStreams.Ole.LocalCheckUpdateTime = new FILETIME();
            _oleDataStreams.Ole.RemoteUpdateTime = new FILETIME();
        }

        private void CreateOleDataStream()
        {
            byte[] oleBytes = ConcatenateByteArrays(BitConverter.GetBytes(_oleDataStreams.Ole.Version),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.Flags),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.LinkUpdateOption),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.Reserved1),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStreamSize),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStream.ClsId.Data1),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStream.ClsId.Data2),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStream.ClsId.Data3),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStream.ClsId.Data4),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStream.StreamData1),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStream.StreamData2),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.ReservedMonikerStream.StreamData3),
                                                    BinaryHelper.GetByteArray( _oleDataStreams.Ole.ReservedMonikerStream.StreamData4, _oleDataStreams.Ole.ReservedMonikerStream.Encoding),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStreamSize),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStream.ClsId.Data1),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStream.ClsId.Data2),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStream.ClsId.Data3),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStream.ClsId.Data4),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStream.StreamData1),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStream.StreamData2),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.RelativeSourceMonikerStream.StreamData3),
                                                    BinaryHelper.GetByteArray(_oleDataStreams.Ole.RelativeSourceMonikerStream.StreamData4, _oleDataStreams.Ole.RelativeSourceMonikerStream.Encoding),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStreamSize),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.ClsId.Data1),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.ClsId.Data2),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.ClsId.Data3),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.ClsId.Data4),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.StreamData1),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.StreamData2),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.StreamData3),
                                                    BinaryHelper.GetByteArray(_oleDataStreams.Ole.AbsoluteSourceMonikerStream.StreamData4, _oleDataStreams.Ole.AbsoluteSourceMonikerStream.Encoding),
                                                    new byte[_oleDataStreams.Ole.ClsIdIndicator],
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.ClsId.Data1),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.ClsId.Data2),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.ClsId.Data3),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.ClsId.Data4),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.ReservedDisplayName.Length),
                                                    BinaryHelper.GetByteArray(_oleDataStreams.Ole.ReservedDisplayName.String, _oleDataStreams.Ole.ReservedDisplayName.Encoding),
                                                    new byte[_oleDataStreams.Ole.Reserved2],
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.LocalUpdateTime.dwLowDateTime),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.LocalUpdateTime.dwHighDateTime),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.LocalCheckUpdateTime.dwLowDateTime),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.LocalCheckUpdateTime.dwHighDateTime),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.RemoteUpdateTime.dwLowDateTime),
                                                    BitConverter.GetBytes(_oleDataStreams.Ole.RemoteUpdateTime.dwHighDateTime));
            _document.Storage.DataStreams.Add("\u0001Ole", oleBytes);
        }

        private void CreateCompObjObject(string AnsiUserTypeString, string Reserved1String)
        {
            _oleDataStreams.CompObj = new CompObjStream();
            _oleDataStreams.CompObj.Header = new CompObjHeader();
            _oleDataStreams.CompObj.AnsiUserType = new LengthPrefixedAnsiString(AnsiUserTypeString);
            _oleDataStreams.CompObj.AnsiClipboardFormat = new ClipboardFormatOrAnsiString();
            _oleDataStreams.CompObj.Reserved1 = new LengthPrefixedAnsiString(Reserved1String);
            _oleDataStreams.CompObj.UnicodeMarker = 0;
            _oleDataStreams.CompObj.UnicodeUserType = new LengthPrefixedUnicodeString();
            _oleDataStreams.CompObj.UnicodeClipboardFormat = new ClipboardFormatOrUnicodeString();
            _oleDataStreams.CompObj.Reserved2 = new LengthPrefixedUnicodeString();
        }

        private void CreateCompObjDataStream()
        {
            byte[] compObjBytes = ConcatenateByteArrays(BitConverter.GetBytes(_oleDataStreams.CompObj.Header.Reserved1),
                                                        BitConverter.GetBytes(_oleDataStreams.CompObj.Header.Version),
                                                        _oleDataStreams.CompObj.Header.Reserved2,
                                                        
                                                        BitConverter.GetBytes(_oleDataStreams.CompObj.AnsiUserType.Length),
                                                        BinaryHelper.GetByteArray(_oleDataStreams.CompObj.AnsiUserType.String+ "\0", _oleDataStreams.CompObj.AnsiUserType.Encoding),
                                                        
                                                        BitConverter.GetBytes(_oleDataStreams.CompObj.AnsiClipboardFormat.MarkerOrLength),
                                                        _oleDataStreams.CompObj.AnsiClipboardFormat.FormatOrAnsiString,
                                                        
                                                        BitConverter.GetBytes(_oleDataStreams.CompObj.Reserved1.Length),
                                                        BinaryHelper.GetByteArray(_oleDataStreams.CompObj.Reserved1.String + "\0", _oleDataStreams.CompObj.Reserved1.Encoding),
                                                        
                                                         BitConverter.GetBytes(_oleDataStreams.CompObj.UnicodeMarker),
                                                        
                                                        BitConverter.GetBytes(_oleDataStreams.CompObj.UnicodeUserType.Length),
                                                        BinaryHelper.GetByteArray(_oleDataStreams.CompObj.UnicodeUserType.String, _oleDataStreams.CompObj.UnicodeUserType.Encoding),
                                                        
                                                        BitConverter.GetBytes(_oleDataStreams.CompObj.UnicodeClipboardFormat.MarkerOrLength),
                                                        _oleDataStreams.CompObj.UnicodeClipboardFormat.FormatOrUnicodeString,

                                                        BitConverter.GetBytes(_oleDataStreams.CompObj.Reserved2.Length),
                                                        BinaryHelper.GetByteArray(_oleDataStreams.CompObj.Reserved2.String, _oleDataStreams.CompObj.Reserved2.Encoding));
            _document.Storage.DataStreams.Add("\u0001CompObj", compObjBytes);
        }

        private void CreateOleNativeObject(byte[] fileData)
        {
            _oleDataStreams.OleNative = new OleNativeStream();
            _oleDataStreams.OleNative.NativeData = fileData;
            _oleDataStreams.OleNative.NativeDataSize = (uint)fileData.Length;
        }

        private void CreateOleNativeDataStream()
        {
            byte[] oleNativeByteSize = BitConverter.GetBytes(_oleDataStreams.OleNative.NativeDataSize);
            byte[] oleNativeBytes = ConcatenateByteArrays(oleNativeByteSize, _oleDataStreams.OleNative.NativeData);
            _document.Storage.DataStreams.Add("\u0001Ole10Native", oleNativeBytes);
        }

        private string GetFileType(string filepath)
        {
            return Path.GetExtension(filepath).ToLower();
        }

        private string EmbedDocument(string filePath)
        {
            string relId = "";
            byte[] fileData = File.ReadAllBytes(filePath);
            string fileType = GetFileType(filePath);
            _oleDataStreams = new OleObjectDataStreams();
            _document = new CompoundDocument();
            if (fileType ==".pdf")
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
                _oleDataStreams.DataFile = fileData;
                _document.Storage.DataStreams.Add("CONTENTS", fileData);
            }
            else if(fileType == ".odt" || fileType == ".ods" || fileType == ".odp") //open office formats
            {
                //Create Ole structure and add data
                CreateOleObject();
                //Create Ole Data Stream and add to Compound object
                CreateOleDataStream();
                //Create CompObj structure and add data
                CreateCompObjObject("OpenDocument Text", "Word.OpenDocumentText.12");
                //Create CompObj Data Stream and add to Compound object
                CreateCompObjDataStream();
                //Add EmbeddedOdf
                _oleDataStreams.DataFile = fileData;
                _document.Storage.DataStreams.Add("EmbeddedOdf", fileData);
            }
            else if(fileType == ".docx" || fileType == ".xlsx" || fileType == ".pptx") //ms office format
            {
                //Embedd as is
                string name = fileType == ".docx" ? "Microsoft_Word_Document" : "";
                name = fileType == ".xlsx" ? "Microsoft_Word_Document" : "";
                name = fileType == ".pptx" ? "Microsoft_Excel_Worksheet" : "";
                int newID = 1;
                var Uri = GetNewUri(_worksheet._package.ZipPackage, "/xl/embeddings/" + name + "{0}" + fileType, ref newID);
                var part = _worksheet._package.ZipPackage.CreatePart(Uri, ContentTypes.contentTypeControlProperties);
                var rel = _worksheet.Part.CreateRelationship(Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/embeddings");
                relId = rel.Id;
                MemoryStream ms = (MemoryStream)part.GetStream(FileMode.Create, FileAccess.Write);
                ms.Write(fileData, 0, fileData.Length);
            }
            else
            {
                //Create CompObj structure and add data
                CreateCompObjObject("OLE Package", "Package");
                //Create CompObj Data Stream and add to Compound object
                CreateCompObjDataStream();
                //Create OleNative structure and add data
                CreateOleNativeObject(fileData);
                //Create OleNative Data Stream and add to Compound object
                CreateOleNativeDataStream();
            }
            if (_document.Storage.DataStreams != null)
            {
                int newID = 1;
                var Uri = GetNewUri(_worksheet._package.ZipPackage, "/xl/embeddings/oleObject{0}.bin", ref newID);
                var part = _worksheet._package.ZipPackage.CreatePart(Uri, ContentTypes.contentTypeOleObject);
                var rel = _worksheet.Part.CreateRelationship(Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/oleObject");
                relId = rel.Id;
                MemoryStream ms = (MemoryStream)part.GetStream(FileMode.Create, FileAccess.Write);
                _document.Save(ms);
            }
            return relId;
        }
        #endregion

        #region OlePres
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

        //private void ReadDEVMODEA(BinaryReader br, ExcelWorksheet ws, ref int ci)
        //{
        //    var dmDeviceName = br.ReadBytes(32);
        //    var dmFormName = br.ReadBytes(32);
        //    var dmSpecVersion = br.ReadUInt16();
        //    var dmDriverVersion = br.ReadUInt16();
        //    var dmSize = br.ReadUInt16();
        //    var dmDriverExtra = br.ReadUInt16();
        //    var dmFields = br.ReadUInt32();
        //    var dmOrientation = br.ReadUInt16();
        //    var dmPaperSize = br.ReadUInt16();
        //    var dmPaperLength = br.ReadUInt16();
        //    var dmPaperWidth = br.ReadUInt16();
        //    var dmScale = br.ReadUInt16();
        //    var dmCopies = br.ReadUInt16();
        //    var dmDefaultSource = br.ReadUInt16();
        //    var dmPrintQuality = br.ReadUInt16();
        //    var dmColor = br.ReadUInt16();
        //    var dmDuplex = br.ReadUInt16();
        //    var dmYResolution = br.ReadUInt16();
        //    var dmTTOption = br.ReadUInt16();
        //    var dmCollate = br.ReadUInt16();
        //    var reserved0 = br.ReadUInt32();
        //    var reserved1 = br.ReadUInt32();
        //    var reserved2 = br.ReadUInt32();
        //    var reserved3 = br.ReadUInt32();
        //    var dmNup = br.ReadUInt32();
        //    var reserved4 = br.ReadUInt32();
        //    var dmICMMethod = br.ReadUInt32();
        //    var dmICMIntent = br.ReadUInt32();
        //    var dmMediaType = br.ReadUInt32();
        //    var dmDitherType = br.ReadUInt32();
        //    var reserved5 = br.ReadUInt32();
        //    var reserved6 = br.ReadUInt32();
        //    var reserved7 = br.ReadUInt32();
        //    var reserved8 = br.ReadUInt32();

        //    ws.Cells[2, ci++].Value = dmDeviceName;
        //    ws.Cells[2, ci++].Value = dmFormName;
        //    ws.Cells[2, ci++].Value = dmSpecVersion;
        //    ws.Cells[2, ci++].Value = dmDriverVersion;
        //    ws.Cells[2, ci++].Value = dmSize;
        //    ws.Cells[2, ci++].Value = dmDriverExtra;
        //    ws.Cells[2, ci++].Value = dmFields;
        //    ws.Cells[2, ci++].Value = dmOrientation;
        //    ws.Cells[2, ci++].Value = dmPaperSize;
        //    ws.Cells[2, ci++].Value = dmPaperLength;
        //    ws.Cells[2, ci++].Value = dmPaperWidth;
        //    ws.Cells[2, ci++].Value = dmScale;
        //    ws.Cells[2, ci++].Value = dmCopies;
        //    ws.Cells[2, ci++].Value = dmDefaultSource;
        //    ws.Cells[2, ci++].Value = dmPrintQuality;
        //    ws.Cells[2, ci++].Value = dmColor;
        //    ws.Cells[2, ci++].Value = dmDuplex;
        //    ws.Cells[2, ci++].Value = dmYResolution;
        //    ws.Cells[2, ci++].Value = dmTTOption;
        //    ws.Cells[2, ci++].Value = dmCollate;
        //    ws.Cells[2, ci++].Value = reserved0;
        //    ws.Cells[2, ci++].Value = reserved1;
        //    ws.Cells[2, ci++].Value = reserved2;
        //    ws.Cells[2, ci++].Value = reserved3;
        //    ws.Cells[2, ci++].Value = dmNup;
        //    ws.Cells[2, ci++].Value = reserved4;
        //    ws.Cells[2, ci++].Value = dmICMMethod;
        //    ws.Cells[2, ci++].Value = dmICMIntent;
        //    ws.Cells[2, ci++].Value = dmMediaType;
        //    ws.Cells[2, ci++].Value = dmDitherType;
        //    ws.Cells[2, ci++].Value = reserved5;
        //    ws.Cells[2, ci++].Value = reserved6;
        //    ws.Cells[2, ci++].Value = reserved7;
        //    ws.Cells[2, ci++].Value = reserved8;
        //}


        //static ushort MinOffset(ushort[] offsets, ushort currentOffset)
        //{
        //    ushort minOffset = ushort.MaxValue;
        //    foreach (ushort offset in offsets)
        //    {
        //        if (offset > currentOffset && offset < minOffset)
        //        {
        //            minOffset = offset;
        //        }
        //    }
        //    return minOffset;
        //}

        //private void ReadDVTARGETDEVICE(BinaryReader br, uint size, ExcelWorksheet ws, ref int ci)
        //{
        //    var DriverNameOffSet = br.ReadUInt16();
        //    var DeviceNameOffSet = br.ReadUInt16();
        //    var PortNameOffSet = br.ReadUInt16();
        //    var ExtDevModeOffSet = br.ReadUInt16();

        //    ws.Cells[2, ci++].Value = DriverNameOffSet;
        //    ws.Cells[2, ci++].Value = DeviceNameOffSet;
        //    ws.Cells[2, ci++].Value = PortNameOffSet;
        //    ws.Cells[2, ci++].Value = ExtDevModeOffSet;

        //    string DriverName = "";
        //    if (DriverNameOffSet != 0)
        //    {
        //        ushort nextOffset = MinOffset(new ushort[] { DeviceNameOffSet, PortNameOffSet, ExtDevModeOffSet, (ushort)size }, DriverNameOffSet);
        //        var DriverNameLength = nextOffset - DriverNameOffSet;
        //        DriverName = BinaryHelper.GetString(br, (uint)DriverNameLength, Encoding.ASCII);
        //    }

        //    ws.Cells[2, ci++].Value = DriverName;

        //    string DeviceName = "";

        //    if (DeviceNameOffSet != 0)
        //    {
        //        ushort nextOffset = MinOffset(new ushort[] { DriverNameOffSet, PortNameOffSet, ExtDevModeOffSet, (ushort)size }, DeviceNameOffSet);
        //        var DeviceNameLength = nextOffset - DeviceNameOffSet;
        //        DeviceName = BinaryHelper.GetString(br, (uint)DeviceNameLength, Encoding.ASCII);
        //    }

        //    ws.Cells[2, ci++].Value = DeviceName;

        //    string PortName = "";
        //    if (PortNameOffSet != 0)
        //    {
        //        ushort nextOffset = MinOffset(new ushort[] { DriverNameOffSet, DeviceNameOffSet, ExtDevModeOffSet, (ushort)size }, PortNameOffSet);
        //        var PortNameLength = nextOffset - PortNameOffSet;
        //        PortName = BinaryHelper.GetString(br, (uint)PortNameLength, Encoding.ASCII);
        //    }

        //    ws.Cells[2, ci++].Value = PortName;

        //    if (ExtDevModeOffSet != 0)
        //        ReadDEVMODEA(br, ws, ref ci); //ExtDevMode
        //    else
        //        ci += 34;
        //}
        #endregion


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

/*2024-08-15
 * 
 * När vi öppnar result filen i excel så blir resultatet att worksheet.xnl är trasig.
 * Den tar bort properties från oleObject taggen
 * progId och r:id tas bort
 * shapeId blir något helt annat. I vår version är shapeId 1025, medan excel gör om den till 2
 * Den tar även helt bort fallback
 * 
 * 
 */


/*
 * TODO:
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