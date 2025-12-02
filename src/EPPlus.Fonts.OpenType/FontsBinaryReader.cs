/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using System;
using System.IO;

namespace EPPlus.Fonts.OpenType
{
    internal class FontsBinaryReader : BinaryReader
    {
        public FontsBinaryReader(Stream input) : base(input)
        {
        }


        private string _context;
        private int _numberOfReadBytes = 0;


        internal void SetContext(string name)
        {
            _context = name;
            _numberOfReadBytes = 0;
        }

        public override byte[] ReadBytes(int count)
        {
            var b = base.ReadBytes(count);
            _numberOfReadBytes += b.Length;
            return b;
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

        internal uint ReadUInt24BigEndian()
        {
            var b = ReadBytes(3);
            return (uint)((b[0] << 16) | (b[1] << 8) | b[2]);
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

        internal long ReadInt64BigEndian()
        {
            var b = ReadBytes(8);
            return BitConverter.ToInt64(new byte[] { b[7], b[6], b[5], b[4], b[3], b[2], b[1], b[0] }, 0);
        }


        internal ushort[] ReadUInt16ArrayBigEndian(int count)
        {
            var result = new ushort[count];
            for (int i = 0; i < count; i++)
            {
                result[i] = ReadUInt16BigEndian();
            }
            return result;
        }

        internal short[] ReadInt16ArrayBigEndian(int count)
        {
            var result = new short[count];
            for (int i = 0; i < count; i++)
            {
                result[i] = ReadInt16BigEndian();
            }
            return result;
        }

    }
}
