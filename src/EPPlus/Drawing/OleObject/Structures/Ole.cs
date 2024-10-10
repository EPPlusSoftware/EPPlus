using OfficeOpenXml.Utils.CompundDocument;
using OfficeOpenXml.Utils;
using System;
using static OfficeOpenXml.Drawing.OleObject.Structures.OleObjectDataStructures;
using System.IO;

namespace OfficeOpenXml.Drawing.OleObject.Structures
{
    internal static class Ole
    {
        internal const string OLE_STREAM_NAME = "\u0001Ole";

        internal static void CreateOleDataStream(OleObjectDataStructures _oleDataStructures, CompoundDocument _document, bool IsExternalLink)
        {
            byte[] oleBytes = BinaryHelper.ConcatenateByteArrays(
                                           BitConverter.GetBytes(_oleDataStructures.Ole.Version),
                                           BitConverter.GetBytes(_oleDataStructures.Ole.Flags),
                                           BitConverter.GetBytes(_oleDataStructures.Ole.LinkUpdateOption),
                                           BitConverter.GetBytes(_oleDataStructures.Ole.Reserved1),
                                           BitConverter.GetBytes(_oleDataStructures.Ole.ReservedMonikerStreamSize));
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
                                        BinaryHelper.GetByteArray(_oleDataStructures.Ole.ReservedMonikerStream.StreamData4, _oleDataStructures.Ole.ReservedMonikerStream.Encoding));
            }
            if (IsExternalLink)
            {
                oleBytes = BinaryHelper.ConcatenateByteArrays(oleBytes,
                                        BitConverter.GetBytes(_oleDataStructures.Ole.RelativeSourceMonikerStreamSize));
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
                                            BinaryHelper.GetByteArray(_oleDataStructures.Ole.RelativeSourceMonikerStream.StreamData4, _oleDataStructures.Ole.RelativeSourceMonikerStream.Encoding));
                }
                oleBytes = BinaryHelper.ConcatenateByteArrays(oleBytes,
                                        BitConverter.GetBytes(_oleDataStructures.Ole.AbsoluteSourceMonikerStreamSize));
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
                                            BinaryHelper.GetByteArray(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.StreamData4, _oleDataStructures.Ole.AbsoluteSourceMonikerStream.Encoding));
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
                                        BitConverter.GetBytes(_oleDataStructures.Ole.RemoteUpdateTime.dwHighDateTime));
            }
            _document.Storage.DataStreams.Add(OLE_STREAM_NAME, new CompoundDocumentItem(OLE_STREAM_NAME, oleBytes));
        }

        internal static void CreateOleObject(OleObjectDataStructures _oleDataStructures, bool IsExternalLink)
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
                                       BinaryHelper.GetByteArray(_oleDataStructures.Ole.ReservedMonikerStream.StreamData4, _oleDataStructures.Ole.ReservedMonikerStream.Encoding));
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
                                    BinaryHelper.GetByteArray(_oleDataStructures.Ole.RelativeSourceMonikerStream.StreamData4, _oleDataStructures.Ole.RelativeSourceMonikerStream.Encoding));
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
                                    BinaryHelper.GetByteArray(_oleDataStructures.Ole.AbsoluteSourceMonikerStream.StreamData4, _oleDataStructures.Ole.AbsoluteSourceMonikerStream.Encoding));
                _oleDataStructures.Ole.AbsoluteSourceMonikerStreamSize = (UInt32)size.Length;
                _oleDataStructures.Ole.ClsId = new CLSID();
                _oleDataStructures.Ole.ReservedDisplayName = new LengthPrefixedUnicodeString();
                _oleDataStructures.Ole.LocalUpdateTime = new FILETIME();
                _oleDataStructures.Ole.LocalCheckUpdateTime = new FILETIME();
                _oleDataStructures.Ole.RemoteUpdateTime = new FILETIME();
            }
        }

        internal static void ReadOleStream(OleObjectDataStructures _oleDataStructures, byte[] oleBytes)
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
                    _oleDataStructures.Ole.ReservedMonikerStream = OleSharedStructures.ReadMONIKERSTREAM(br, _oleDataStructures.Ole.ReservedMonikerStreamSize - 4);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.Ole.RelativeSourceMonikerStreamSize = br.ReadUInt32();
                if (_oleDataStructures.Ole.RelativeSourceMonikerStreamSize != 0)
                    _oleDataStructures.Ole.RelativeSourceMonikerStream = OleSharedStructures.ReadMONIKERSTREAM(br, _oleDataStructures.Ole.RelativeSourceMonikerStreamSize - 4);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.Ole.AbsoluteSourceMonikerStreamSize = br.ReadUInt32();
                if (_oleDataStructures.Ole.AbsoluteSourceMonikerStreamSize != 0)
                    _oleDataStructures.Ole.AbsoluteSourceMonikerStream = OleSharedStructures.ReadMONIKERSTREAM(br, _oleDataStructures.Ole.AbsoluteSourceMonikerStreamSize - 4);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.Ole.ClsIdIndicator = br.ReadUInt32();
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.Ole.ClsId = OleSharedStructures.ReadCLSID(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.Ole.ReservedDisplayName = OleSharedStructures.ReadLengthPrefixedUnicodeString(br);
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.Ole.Reserved2 = br.ReadUInt32();
                if (br.BaseStream.Position >= br.BaseStream.Length)
                    return;

                _oleDataStructures.Ole.LocalUpdateTime = OleSharedStructures.ReadFILETIME(br);
                _oleDataStructures.Ole.LocalCheckUpdateTime = OleSharedStructures.ReadFILETIME(br);
                _oleDataStructures.Ole.RemoteUpdateTime = OleSharedStructures.ReadFILETIME(br);
            }
        }
    }
}
