using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeOpenXml.Utils;
using static OfficeOpenXml.Drawing.OleObject.Structures.OleObjectDataStructures;

namespace OfficeOpenXml.Drawing.OleObject.Structures
{
    internal static class OleSharedStructures
    {
        internal static LengthPrefixedUnicodeString ReadLengthPrefixedUnicodeString(BinaryReader br)
        {
            LengthPrefixedUnicodeString LPUniS = new LengthPrefixedUnicodeString();
            LPUniS.Length = br.ReadUInt32();
            LPUniS.String = BinaryHelper.GetString(br, LPUniS.Length * 2, LPUniS.Encoding);
            return LPUniS;
        }
        internal static LengthPrefixedAnsiString ReadLengthPrefixedAnsiString(BinaryReader br)
        {
            LengthPrefixedAnsiString LPAnsiS = new LengthPrefixedAnsiString();
            LPAnsiS.Length = br.ReadUInt32();
            LPAnsiS.String = BinaryHelper.GetString(br, LPAnsiS.Length, LPAnsiS.Encoding).Trim('\0');
            return LPAnsiS;
        }
        internal static LengthPrefixedAnsiString ReadUntilNullTerminator(BinaryReader br)
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
        internal static ClipboardFormatOrUnicodeString ReadClipboardFormatOrUnicodeString(BinaryReader br)
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
        internal static ClipboardFormatOrAnsiString ReadClipboardFormatOrAnsiString(BinaryReader br)
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
        internal static CLSID ReadCLSID(BinaryReader br)
        {
            CLSID CLSID = new CLSID();
            CLSID.Data1 = br.ReadUInt32();
            CLSID.Data2 = br.ReadUInt16();
            CLSID.Data3 = br.ReadUInt16();
            CLSID.Data4 = br.ReadUInt64();
            return CLSID;
        }
        internal static MonikerStream ReadMONIKERSTREAM(BinaryReader br, uint size)
        {
            MonikerStream monikerStream = new MonikerStream();
            monikerStream.ClsId = ReadCLSID(br);
            monikerStream.StreamData1 = br.ReadUInt32();
            monikerStream.StreamData2 = br.ReadUInt16();
            monikerStream.StreamData3 = br.ReadUInt32();
            monikerStream.StreamData4 = BinaryHelper.GetString(br, monikerStream.StreamData3, Encoding.ASCII);
            return monikerStream;
        }
        internal static FILETIME ReadFILETIME(BinaryReader br)
        {
            FILETIME FILETIME = new FILETIME();
            FILETIME.dwLowDateTime = br.ReadUInt32();
            FILETIME.dwHighDateTime = br.ReadUInt32();
            return FILETIME;
        }
    }
}
