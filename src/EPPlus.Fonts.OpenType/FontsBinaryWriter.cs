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

        private int _bytesWritten = 0;
        private bool _hit = false;

        private void IncreaseBytesWritten(int nBytes)
        {
            _bytesWritten += nBytes;
            if(_bytesWritten >= 53 && !_hit)
            {
                _hit = true;
            }
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


        internal void WriteUInt24BigEndian(uint value)
        {
            if (value > 0xFFFFFF)
                throw new ArgumentOutOfRangeException(nameof(value), "Value exceeds 24-bit range.");

            var bytes = BitConverter.GetBytes(value);
            Write(bytes[2]); // high byte
            Write(bytes[1]); // mid byte
            Write(bytes[0]); // low byte
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

        internal void WriteInt64BigEndian(long value)
        {
            var bytes = BitConverter.GetBytes(value);
            Write(bytes[7]);
            Write(bytes[6]);
            Write(bytes[5]);
            Write(bytes[4]);
            Write(bytes[3]);
            Write(bytes[2]);
            Write(bytes[1]);
            Write(bytes[0]);
        }

        public override void Write(byte[] buffer)
        {
            base.Write(buffer);
            IncreaseBytesWritten(buffer.Length);
        }

        public override void Write(byte value)
        {
            base.Write(value);
            IncreaseBytesWritten(1);
        }

        public override void Write(int i)
        {
            base.Write(i);
            IncreaseBytesWritten(4);
        }

        public override void Write(uint i)
        {
            base.Write(i);
            IncreaseBytesWritten(4);
        }

        public override void Write(ushort value)
        {
            base.Write(value);
            IncreaseBytesWritten(2);
        }

        public override void Write(short value)
        {
            base.Write(value);
            IncreaseBytesWritten(2);
        }

    }
}
