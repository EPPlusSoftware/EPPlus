/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts
{
    internal class BigEndianBinaryReader : BinaryReader
    {
        internal BigEndianBinaryReader(Stream input) : base(input)
        {
        }

        public ushort ReadUInt16BigEndian()
        {
            var b = ReadBytes(2);
            return BitConverter.ToUInt16(new byte[] { b[1], b[0] }, 0);
        }
        public short ReadInt16BigEndian()
        {
            var b = ReadBytes(2);
            return BitConverter.ToInt16(new byte[] { b[1], b[0] }, 0);
        }
        public int ReadInt32BigEndian()
        {
            var b = ReadBytes(4);
            return BitConverter.ToInt32(new byte[] { b[3], b[2], b[1], b[0] }, 0);
        }

        public uint ReadUInt32BigEndian()
        {
            var b = ReadBytes(4);
            return BitConverter.ToUInt32(new byte[] { b[3], b[2], b[1], b[0] }, 0);
        }

    }
}
