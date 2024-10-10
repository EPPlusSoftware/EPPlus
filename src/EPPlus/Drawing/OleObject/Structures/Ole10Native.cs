using OfficeOpenXml.Utils.CompundDocument;
using System;
using System.IO;
using OfficeOpenXml.Utils;
using static OfficeOpenXml.Drawing.OleObject.Structures.OleObjectDataStructures;

namespace OfficeOpenXml.Drawing.OleObject.Structures
{
    internal static class Ole10Native
    {
        internal const string OLE10NATIVE_STREAM_NAME = "\u0001Ole10Native";

        internal static void CreateOle10NativeDataStream(OleObjectDataStructures _oleDataStructures, CompoundDocument _document)
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
                                                 BinaryHelper.GetByteArray(_oleDataStructures.OleNative.Footer.FilePath.String, _oleDataStructures.OleNative.Footer.FilePath.Encoding));
            //Write total size to size.
            var totalsize = BitConverter.GetBytes(oleNativeBytes.Length - 4);
            oleNativeBytes[0] = totalsize[0];
            oleNativeBytes[1] = totalsize[1];
            oleNativeBytes[2] = totalsize[2];
            oleNativeBytes[3] = totalsize[3];
            _document.Storage.DataStreams.Add(OLE10NATIVE_STREAM_NAME, new CompoundDocumentItem(OLE10NATIVE_STREAM_NAME, oleNativeBytes));
        }

        internal static void CreateOle10NativeObject(byte[] fileData, string filePath, OleObjectDataStructures _oleDataStructures)
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

        private static Ole10NativeHeader ReadOle10NativeHeader(BinaryReader br)
        {
            Ole10NativeHeader header = new Ole10NativeHeader();
            header.Size = br.ReadUInt32();
            header.Type = br.ReadUInt16();
            header.FileName = OleSharedStructures.ReadUntilNullTerminator(br);
            header.FilePath = OleSharedStructures.ReadUntilNullTerminator(br);
            header.Reserved1 = br.ReadUInt32();
            header.TempPath = OleSharedStructures.ReadLengthPrefixedAnsiString(br);
            return header;
        }

        private static Ole10NativeFooter ReadOle10NativeFooter(BinaryReader br)
        {
            Ole10NativeFooter footer = new Ole10NativeFooter();
            footer.TempPath = OleSharedStructures.ReadLengthPrefixedUnicodeString(br);
            footer.FileName = OleSharedStructures.ReadLengthPrefixedUnicodeString(br);
            footer.FilePath = OleSharedStructures.ReadLengthPrefixedUnicodeString(br);
            return footer;
        }

        internal static void ReadOle10Native(OleObjectDataStructures _oleDataStructures, byte[] oleBytes)
        {
            using (var ms = RecyclableMemory.GetStream(oleBytes))
            {
                BinaryReader br = new BinaryReader(ms);
                _oleDataStructures.OleNative.Header = ReadOle10NativeHeader(br);
                _oleDataStructures.OleNative.NativeDataSize = br.ReadUInt32();
                _oleDataStructures.OleNative.NativeData = br.ReadBytes((int)_oleDataStructures.OleNative.NativeDataSize);
                _oleDataStructures.OleNative.Footer = ReadOle10NativeFooter(br);
            }
        }
    }
}
