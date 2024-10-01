using OfficeOpenXml.Utils.CompundDocument;
using OfficeOpenXml.Utils;
using System;
using static OfficeOpenXml.Drawing.OleObject.Structures.OleObjectDataStructures;
using System.IO;

namespace OfficeOpenXml.Drawing.OleObject.Structures
{
    internal static class CompObj
    {
        internal const string COMPOBJ_STREAM_NAME = "\u0001CompObj";

        internal static void CreateCompObjDataStream(OleObjectDataStructures _oleDataStructures, CompoundDocument _document)
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
                                               BinaryHelper.GetByteArray(_oleDataStructures.CompObj.Reserved2.String, _oleDataStructures.CompObj.Reserved2.Encoding));
            _document.Storage.DataStreams.Add(COMPOBJ_STREAM_NAME, new CompoundDocumentItem(COMPOBJ_STREAM_NAME, compObjBytes));
        }

        internal static void CreateCompObjObject(OleObjectDataStructures _oleDataStructures, string AnsiUserTypeString, string Reserved1String)
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

        private static CompObjHeader ReadCompObjHeader(BinaryReader br)
        {
            CompObjHeader header = new CompObjHeader();
            header.Reserved1 = br.ReadUInt32();
            header.Version = br.ReadUInt32();
            header.Reserved2 = br.ReadBytes(20);
            return header;
        }
        internal static void ReadCompObjStream(OleObjectDataStructures _oleDataStructures, byte[] oleBytes)
        {
            using (var ms = RecyclableMemory.GetStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);
                _oleDataStructures.CompObj.Header = ReadCompObjHeader(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.CompObj.AnsiUserType = OleSharedStructures.ReadLengthPrefixedAnsiString(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.CompObj.AnsiClipboardFormat = OleSharedStructures.ReadClipboardFormatOrAnsiString(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.CompObj.Reserved1 = OleSharedStructures.ReadLengthPrefixedAnsiString(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.CompObj.UnicodeMarker = br.ReadUInt32();
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.CompObj.UnicodeUserType = OleSharedStructures.ReadLengthPrefixedUnicodeString(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.CompObj.UnicodeClipboardFormat = OleSharedStructures.ReadClipboardFormatOrUnicodeString(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.CompObj.Reserved2 = OleSharedStructures.ReadLengthPrefixedUnicodeString(br);
            }
        }
    }
}
