using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType
{
    internal class FontsBinaryWriter : BinaryWriter
    {

        internal FontsBinaryWriter(Stream output) : base(output)
        {
        }

        internal void WriteUInt16BigEndian(ushort value)
        {
            var bytes = BitConverter.GetBytes(value);
            Write(bytes[1]); // MSB
            Write(bytes[0]); // LSB
        }

        internal void WriteInt16BigEndian(short value)
        {
            var bytes = BitConverter.GetBytes(value);
            Write(bytes[1]);
            Write(bytes[0]);
        }

        internal void WriteUInt32BigEndian(uint value)
        {
            var bytes = BitConverter.GetBytes(value);
            Write(bytes[3]);
            Write(bytes[2]);
            Write(bytes[1]);
            Write(bytes[0]);
        }

        internal void WriteInt32BigEndian(int value)
        {
            var bytes = BitConverter.GetBytes(value);
            Write(bytes[3]);
            Write(bytes[2]);
            Write(bytes[1]);
            Write(bytes[0]);
        }

    }
}
