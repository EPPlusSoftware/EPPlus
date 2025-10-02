using System;
using System.IO;

namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader
{
    internal class MyBinaryReader : BinaryReader
    {
        public MyBinaryReader(Stream input) : base(input)
        {
        }

        internal ushort ReadUInt16BigEndian()
        {
            var b = ReadBytes(2);
            return BitConverter.ToUInt16(new byte[] { b[1], b[0] }, 0);
        }
        internal short ReadInt16BigEndian()
        {
            var b = ReadBytes(2);
            return BitConverter.ToInt16(new byte[] { b[1], b[0] }, 0);
        }
        internal int ReadInt32BigEndian()
        {
            var b = ReadBytes(4);
            return BitConverter.ToInt32(new byte[] { b[3], b[2], b[1], b[0] }, 0);
        }

        internal uint ReadUInt32BigEndian()
        {
            var b = ReadBytes(4);
            return BitConverter.ToUInt32(new byte[] { b[3], b[2], b[1], b[0] }, 0);
        }

    }
}
