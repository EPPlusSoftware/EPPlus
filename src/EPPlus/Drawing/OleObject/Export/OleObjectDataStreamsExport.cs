using OfficeOpenXml.Drawing.OleObject.Structures;

namespace OfficeOpenXml.Drawing.OleObject
{
    internal static class OleObjectDataStreamsExport
    {

        private static void ExportLengthPrefixedUnicodeString(ExcelWorksheet ws, ref int c, OleObjectDataStructures.LengthPrefixedUnicodeString LPUniS)
        {
            if (LPUniS == null)
            {
                c += 2;
                return;
            }
            ws.Cells[2, c++].Value = LPUniS.Length;
            ws.Cells[2, c++].Value = LPUniS.String;
        }

        private static void ExportLengthPrefixedAnsiString(ExcelWorksheet ws, ref int c, OleObjectDataStructures.LengthPrefixedAnsiString LPAnsiS)
        {
            if (LPAnsiS == null)
            {
                c += 2;
                return;
            }
            ws.Cells[2, c++].Value = LPAnsiS.Length;
            ws.Cells[2, c++].Value = LPAnsiS.String;
        }

        private static void ExportClipboardFormatOrUnicodeString(ExcelWorksheet ws, ref int c, OleObjectDataStructures.ClipboardFormatOrUnicodeString CFOUS)
        {
            if (CFOUS == null)
            {
                c += 2;
                return;
            }
            ws.Cells[2, c++].Value = CFOUS.MarkerOrLength;
            ws.Cells[2, c++].Value = CFOUS.FormatOrUnicodeString;
        }

        private static void ExportClipboardFormatOrAnsiString(ExcelWorksheet ws, ref int c, OleObjectDataStructures.ClipboardFormatOrAnsiString CFOAS)
        {
            if (CFOAS == null)
            {
                c += 2;
                return;
            }
            ws.Cells[2, c++].Value = CFOAS.MarkerOrLength;
            ws.Cells[2, c++].Value = CFOAS.FormatOrAnsiString;
        }

        private static void ExportCLSID(ExcelWorksheet ws, ref int c, OleObjectDataStructures.CLSID ClsId)
        {
            if (ClsId == null)
            {
                c += 4;
                return;
            }
            ws.Cells[2, c++].Value = ClsId.Data1;
            ws.Cells[2, c++].Value = ClsId.Data2;
            ws.Cells[2, c++].Value = ClsId.Data3;
            ws.Cells[2, c++].Value = ClsId.Data4;
        }

        private static void ExportMonikerStream(ExcelWorksheet ws, ref int c, OleObjectDataStructures.MonikerStream MonikerStream)
        {
            if (MonikerStream == null)
            {
                c += 8;
                return;
            }
            ExportCLSID(ws, ref c, MonikerStream.ClsId);
            ws.Cells[2, c++].Value = MonikerStream.StreamData1;
            ws.Cells[2, c++].Value = MonikerStream.StreamData2;
            ws.Cells[2, c++].Value = MonikerStream.StreamData3;
            ws.Cells[2, c++].Value = MonikerStream.StreamData4;
        }

        private static void ExportFILETIME(ExcelWorksheet ws, ref int c, OleObjectDataStructures.FILETIME FILETIME)
        {
            if (FILETIME == null)
            {
                c += 2;
                return;
            }
            ws.Cells[2, c++].Value = FILETIME.dwLowDateTime;
            ws.Cells[2, c++].Value = FILETIME.dwHighDateTime;
        }

        private static void ExportCompObjHeader(ExcelWorksheet ws, ref int c, OleObjectDataStructures.CompObjHeader header)
        {
            if (header == null)
            {
                c += 3;
                return;
            }
            ws.Cells[2, c++].Value = header.Reserved1;
            ws.Cells[2, c++].Value = header.Version;
            ws.Cells[2, c++].Value = header.Reserved2;
        }

        internal static void ExportCompObj(string currentPackageFileName, string currentPackageOleObjectName, ExcelPackage newPackage, OleObjectDataStructures oleDataStreams)
        {
            if (oleDataStreams.CompObj == null)
                return;

            int c = 1;
            ExcelWorksheet ws = newPackage.Workbook.Worksheets["CompObj"];
            if (ws == null)
            {
                ws = newPackage.Workbook.Worksheets.Add("CompObj");
                ws.Cells[1, c++].Value = "File";
                ws.Cells[1, c++].Value = "OleObject File";
                ws.Cells[1, c++].Value = "OleObject";
                ws.Cells[1, c++].Value = "CompObjHeader Reserved1";
                ws.Cells[1, c++].Value = "CompObjHeader Version";
                ws.Cells[1, c++].Value = "CompObjHeader Reserved2";
                ws.Cells[1, c++].Value = "AnsiUserType Length";
                ws.Cells[1, c++].Value = "AnsiUserType String";
                ws.Cells[1, c++].Value = "AnsiClipboardFormat MarkerOrLength";
                ws.Cells[1, c++].Value = "AnsiClipboardFormat FormatOrAnsiString";
                ws.Cells[1, c++].Value = "Reserved1 Length";
                ws.Cells[1, c++].Value = "Reserved1 String";
                ws.Cells[1, c++].Value = "UnicodeMarker";
                ws.Cells[1, c++].Value = "UnicodeUserType Length";
                ws.Cells[1, c++].Value = "UnicodeUserType String";
                ws.Cells[1, c++].Value = "UnicodeClipboardFormat MarkerOrLength";
                ws.Cells[1, c++].Value = "UnicodeClipboardFormat FormatOrUniodeString";
                ws.Cells[1, c++].Value = "Reserved2 Length";
                ws.Cells[1, c++].Value = "Reserved2 String";

            }
            ws.InsertRow(2, 1);
            c = 1;
            ws.Cells[2, c++].Value = currentPackageFileName;
            ws.Cells[2, c++].Value = currentPackageOleObjectName;
            ExportCompObjHeader(ws, ref c, oleDataStreams.CompObj.Header);
            ExportLengthPrefixedAnsiString(ws, ref c, oleDataStreams.CompObj.AnsiUserType);
            ExportClipboardFormatOrAnsiString(ws, ref c, oleDataStreams.CompObj.AnsiClipboardFormat);
            ExportLengthPrefixedAnsiString(ws, ref c, oleDataStreams.CompObj.Reserved1);
            ws.Cells[2, c++].Value = oleDataStreams.CompObj.UnicodeMarker;
            ExportLengthPrefixedUnicodeString(ws, ref c, oleDataStreams.CompObj.UnicodeUserType);
            ExportClipboardFormatOrUnicodeString(ws, ref c, oleDataStreams.CompObj.UnicodeClipboardFormat);
            ExportLengthPrefixedUnicodeString(ws, ref c, oleDataStreams.CompObj.Reserved2);
        }

        internal static void ExportOle(string currentPackageFileName, string currentPackageOleObjectName, ExcelPackage newPackage, OleObjectDataStructures oleDataStreams, bool isExternalLink = false)
        {
            if (oleDataStreams.Ole == null)
                return;

            int c = 1;
            ExcelWorksheet ws = newPackage.Workbook.Worksheets["Ole"];
            if (ws == null)
            {
                ws = newPackage.Workbook.Worksheets.Add("Ole");
                ws.Cells[1, c++].Value = "File";
                ws.Cells[1, c++].Value = "OleObject File";
                ws.Cells[1, c++].Value = "Version";
                ws.Cells[1, c++].Value = "Flags";
                ws.Cells[1, c++].Value = "LinkUpdateOption";
                ws.Cells[1, c++].Value = "Reserved1";
                ws.Cells[1, c++].Value = "ReservedMonikerStreamSize";
                ws.Cells[1, c++].Value = "ReservedMonikerStream ClsId Data 1";
                ws.Cells[1, c++].Value = "ReservedMonikerStream ClsId Data 2";
                ws.Cells[1, c++].Value = "ReservedMonikerStream ClsId Data 3";
                ws.Cells[1, c++].Value = "ReservedMonikerStream ClsId Data 4";
                ws.Cells[1, c++].Value = "ReservedMonikerStream Data 1";
                ws.Cells[1, c++].Value = "ReservedMonikerStream Data 2";
                ws.Cells[1, c++].Value = "ReservedMonikerStream Data 3";
                ws.Cells[1, c++].Value = "ReservedMonikerStream Data 4";
                ws.Cells[1, c++].Value = "RelativeSourceMonikerStreamSize";
                ws.Cells[1, c++].Value = "RelativeSourceMonikerStream ClsId Data 1";
                ws.Cells[1, c++].Value = "RelativeSourceMonikerStream ClsId Data 2";
                ws.Cells[1, c++].Value = "RelativeSourceMonikerStream ClsId Data 3";
                ws.Cells[1, c++].Value = "RelativeSourceMonikerStream ClsId Data 4";
                ws.Cells[1, c++].Value = "RelativeSourceMonikerStream Data 1";
                ws.Cells[1, c++].Value = "RelativeSourceMonikerStream Data 2";
                ws.Cells[1, c++].Value = "RelativeSourceMonikerStream Data 3";
                ws.Cells[1, c++].Value = "RelativeSourceMonikerStream Data 4";
                ws.Cells[1, c++].Value = "AbsoluteSourceMonikerStreamSize";
                ws.Cells[1, c++].Value = "AbsoluteSourceMonikerStream ClsId Data 1";
                ws.Cells[1, c++].Value = "AbsoluteSourceMonikerStream ClsId Data 2";
                ws.Cells[1, c++].Value = "AbsoluteSourceMonikerStream ClsId Data 3";
                ws.Cells[1, c++].Value = "AbsoluteSourceMonikerStream ClsId Data 4";
                ws.Cells[1, c++].Value = "AbsoluteSourceMonikerStream Data 1";
                ws.Cells[1, c++].Value = "AbsoluteSourceMonikerStream Data 2";
                ws.Cells[1, c++].Value = "AbsoluteSourceMonikerStream Data 3";
                ws.Cells[1, c++].Value = "AbsoluteSourceMonikerStream Data 4";
                ws.Cells[1, c++].Value = "ClsIdIndicator";
                ws.Cells[1, c++].Value = "ClsId Data 1";
                ws.Cells[1, c++].Value = "ClsId Data 2";
                ws.Cells[1, c++].Value = "ClsId Data 3";
                ws.Cells[1, c++].Value = "ClsId Data 4";
                ws.Cells[1, c++].Value = "ReservedDisplayName.Length";
                ws.Cells[1, c++].Value = "ReservedDisplayName.String";
                ws.Cells[1, c++].Value = "Reserved2";
                ws.Cells[1, c++].Value = "LocalUpdateTime dwLowDateTime";
                ws.Cells[1, c++].Value = "LocalUpdateTime dwHighDateTime";
                ws.Cells[1, c++].Value = "LocalCheckUpdateTime dwLowDateTime";
                ws.Cells[1, c++].Value = "LocalCheckUpdateTime dwHighDateTime";
                ws.Cells[1, c++].Value = "RemoteUpdateTime dwLowDateTime";
                ws.Cells[1, c++].Value = "RemoteUpdateTime dwHighDateTime";
            }
            ws.InsertRow(2, 1);
            c = 1;
            ws.Cells[2, c++].Value = currentPackageFileName;
            ws.Cells[2, c++].Value = currentPackageOleObjectName;
            ws.Cells[2, c++].Value = oleDataStreams.Ole.Version;
            ws.Cells[2, c++].Value = oleDataStreams.Ole.Flags;
            ws.Cells[2, c++].Value = oleDataStreams.Ole.LinkUpdateOption;
            ws.Cells[2, c++].Value = oleDataStreams.Ole.Reserved1;
            ws.Cells[2, c++].Value = oleDataStreams.Ole.ReservedMonikerStreamSize;
            ExportMonikerStream(ws, ref c, oleDataStreams.Ole.ReservedMonikerStream);
            if (isExternalLink)
            {
                ws.Cells[2, c++].Value = oleDataStreams.Ole.RelativeSourceMonikerStreamSize;
                ExportMonikerStream(ws, ref c, oleDataStreams.Ole.RelativeSourceMonikerStream);
                ws.Cells[2, c++].Value = oleDataStreams.Ole.AbsoluteSourceMonikerStreamSize;
                ExportMonikerStream(ws, ref c, oleDataStreams.Ole.AbsoluteSourceMonikerStream);
                ws.Cells[2, c++].Value = oleDataStreams.Ole.ClsIdIndicator;
                ExportCLSID(ws, ref c, oleDataStreams.Ole.ClsId);
                ExportLengthPrefixedUnicodeString(ws, ref c, oleDataStreams.Ole.ReservedDisplayName);
                ws.Cells[2, c++].Value = oleDataStreams.Ole.Reserved2;
                ExportFILETIME(ws, ref c, oleDataStreams.Ole.LocalUpdateTime);
                ExportFILETIME(ws, ref c, oleDataStreams.Ole.LocalCheckUpdateTime);
                ExportFILETIME(ws, ref c, oleDataStreams.Ole.RemoteUpdateTime);
            }
        }

        internal static void ExportOleNative(string currentPackageFileName, string currentPackageOleObjectName, ExcelPackage newPackage, OleObjectDataStructures oleDataStreams)
        {
            if (oleDataStreams.OleNative == null)
                return;
            int c = 1;
            ExcelWorksheet ws = newPackage.Workbook.Worksheets["Ole10Native"];
            if (ws == null)
            {
                ws = newPackage.Workbook.Worksheets.Add("Ole10Native");
                ws.Cells[1, c++].Value = "File";
                ws.Cells[1, c++].Value = "OleObject File";
                ws.Cells[1, c++].Value = "Size";
                ws.Cells[1, c++].Value = "Type";
                ws.Cells[1, c++].Value = "FileName";
                ws.Cells[1, c++].Value = "FilePath";
                ws.Cells[1, c++].Value = "Reserved1";
                ws.Cells[1, c++].Value = "TempPath.Length";
                ws.Cells[1, c++].Value = "TempPath.String";
                ws.Cells[1, c++].Value = "NativeDataSize";
                ws.Cells[1, c++].Value = "NativeData";
                ws.Cells[1, c++].Value = "TempPath.Length";
                ws.Cells[1, c++].Value = "TempPath.String";
                ws.Cells[1, c++].Value = "FileName.Length";
                ws.Cells[1, c++].Value = "FileName.String";
                ws.Cells[1, c++].Value = "FilePath.Length";
                ws.Cells[1, c++].Value = "FilePath.String";
            }
            ws.InsertRow(2, 1);
            c = 1;
            ws.Cells[2, c++].Value = currentPackageFileName;
            ws.Cells[2, c++].Value = currentPackageOleObjectName;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.Header.Size;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.Header.Type;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.Header.FileName.String;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.Header.FilePath.String;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.Header.Reserved1;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.Header.TempPath.Length;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.Header.TempPath.String;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.NativeDataSize;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.NativeData;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.Footer.TempPath.Length;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.Footer.TempPath.String;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.Footer.FileName.Length;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.Footer.FileName.String;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.Footer.FilePath.Length;
            ws.Cells[2, c++].Value = oleDataStreams.OleNative.Footer.FilePath.String;
        }
    }
}
