using System;
using System.Text;

namespace OfficeOpenXml.Drawing.OleObject
{
    internal class OleObjectDataStreams
    {

        internal OleNativeStream OleNative;
        internal OleObjectStream Ole;
        internal CompObjStream CompObj;
        internal byte[] DataFile;

        internal class MonikerStream
        {
            internal CLSID ClsId;
            internal UInt32 StreamData1 = 2;
            internal UInt16 StreamData2 = 33;
            internal UInt32 StreamData3 = 0; //Size of StreamData4
            internal string StreamData4 = "";
            internal Encoding Encoding = Encoding.Unicode;
        }

        internal class CLSID
        {
            internal UInt32 Data1 = 772;
            internal UInt16 Data2 = 0;
            internal UInt16 Data3 = 0;
            internal UInt64 Data4 = 5044030000000000000;
        }

        internal class LengthPrefixedUnicodeString
        {
            internal UInt32 Length = 0;
            internal string String;
            internal Encoding Encoding = Encoding.Unicode;
            internal LengthPrefixedUnicodeString() { }
            internal LengthPrefixedUnicodeString(string String)
            {
                this.String = String;
                Length = (UInt32)String.Length;
            }
        }

        internal class LengthPrefixedAnsiString
        {
            internal UInt32 Length = 0;
            internal string String;
            internal Encoding Encoding = Encoding.ASCII;
            internal LengthPrefixedAnsiString() { }
            internal LengthPrefixedAnsiString(string String)
            {
                this.String = String;
                Length = (UInt32)String.Length + 1;
            }
        }

        internal class ClipboardFormatOrUnicodeString
        {
            //If this is set to 0x00000000, the FormatOrUnicodeString field MUST
            //NOT be present.If this is set to 0xffffffff or 0xfffffffe, the FormatOrUnicodeString field MUST be
            //4 bytes in size and MUST contain a standard clipboard format identifier
            //Otherwise, the FormatOrUnicodeString field MUST be set to a Unicode string containing the name of a registered clipboard format
            //and the MarkerOrLength field MUST be set to the number of Unicode characters in the FormatOrUnicodeString field, including the
            //terminating null character.
            internal UInt32 MarkerOrLength = 0;
            internal Byte[] FormatOrUnicodeString;
            internal ClipboardFormatOrUnicodeString()
            {
                FormatOrUnicodeString = new byte[0];
            }
            internal ClipboardFormatOrUnicodeString(UInt32 MarkerOrLength, Byte[] FormatOrUnicodeString)
            {
                this.MarkerOrLength = MarkerOrLength;
                if (MarkerOrLength == 0)
                {
                    FormatOrUnicodeString = new byte[0];
                    return;
                }
                else
                {
                    this.FormatOrUnicodeString = FormatOrUnicodeString;
                }
            }
        }

        internal class ClipboardFormatOrAnsiString
        {
            //If this field is set to 0xFFFFFFFF or 0xFFFFFFFE,
            //the FormatOrAnsiString field MUST be 4 bytes in size and MUST contain a standard clipboard format identifier.
            //If this set to a value other than 0x00000000,
            //the FormatOrAnsiString field MUST be set to a null-terminated ANSI string containing the name of a registered clipboard format
            internal UInt32 MarkerOrLength = 0;
            internal Byte[] FormatOrAnsiString;
            internal ClipboardFormatOrAnsiString()
            {
                FormatOrAnsiString = new byte[0];
            }
            internal ClipboardFormatOrAnsiString(UInt32 MarkerOrLength, Byte[] FormatOrAnsiString)
            {
                this.MarkerOrLength = MarkerOrLength;
                if (MarkerOrLength == 0)
                {
                    FormatOrAnsiString = new byte[0];
                    return;
                }
                else
                {
                    this.FormatOrAnsiString = FormatOrAnsiString;
                }
            }
        }

        internal class FILETIME
        {
            internal UInt32 dwLowDateTime;
            internal UInt32 dwHighDateTime;
        }

        internal class OleObjectStream
        {
            internal UInt32 Version = 33554433;
            internal UInt32 Flags = 0;
            internal UInt32 LinkUpdateOption = 0;
            internal UInt32 Reserved1 = 0;
            internal UInt32 ReservedMonikerStreamSize = 0; //Subtract by 4 when reading if not 0
            internal MonikerStream ReservedMonikerStream;

            internal UInt32 RelativeSourceMonikerStreamSize = 0; //Subtract by 4 when reading if not 0
            internal MonikerStream RelativeSourceMonikerStream;

            internal UInt32 AbsoluteSourceMonikerStreamSize = 0; //Subtract by 4 when reading if not 0
            internal MonikerStream AbsoluteSourceMonikerStream;

            internal UInt32 ClsIdIndicator = 0;
            internal CLSID ClsId;

            internal LengthPrefixedUnicodeString ReservedDisplayName;

            internal UInt32 Reserved2 = 0;

            internal FILETIME LocalUpdateTime;
            internal FILETIME LocalCheckUpdateTime;
            internal FILETIME RemoteUpdateTime;
        }

        internal class CompObjHeader
        {
            internal UInt32 Reserved1 = 4294836225;
            internal UInt32 Version = 2563;
            internal byte[] Reserved2 = new byte[20] {255,255,255,255,12,0,3,0,0,0,0,0,192,0,0,0,0,0,0,70 };
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
            internal UInt32 UnicodeMarker = 1907505652; //If this field is present and is NOT set to 0x71B239F4, the remaining fields of the structure MUST be ignored on processing.
            internal LengthPrefixedUnicodeString UnicodeUserType;
            internal ClipboardFormatOrUnicodeString UnicodeClipboardFormat; //MarkerOrLength field of the ClipboardFormatOrUnicodeString structure contains a value other than 0x00000000, 0xffffffff, or 0xfffffffe, the value MUST NOT be more than 0x00000190. Otherwise, the CompObjStream structure is invalid.
            internal LengthPrefixedUnicodeString Reserved2;
        }

        internal class OleNativeStream
        {
            internal UInt32 NativeDataSize = 0;
            internal byte[] NativeData;
        }
    }
}
