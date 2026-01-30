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
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Serialization
{
    internal class CmapSubtable4Deserializer
    {
        public CmapSubtable4Deserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        private readonly FontsBinaryReader _reader;


        public CmapSubtable4 Deserialize(uint startIndex)
        {
            _reader.BaseStream.Seek(startIndex, SeekOrigin.Begin);

            var format = _reader.ReadUInt16BigEndian();

            var table = new CmapSubtable4
            {
                Length = _reader.ReadUInt16BigEndian(),
                Language = _reader.ReadUInt16BigEndian(),
                SegCountX2 = _reader.ReadUInt16BigEndian(),
                SearchRange = _reader.ReadUInt16BigEndian(),
                EntrySelector = _reader.ReadUInt16BigEndian(),
                RangeShift = _reader.ReadUInt16BigEndian()
            };

            int segCount = table.SegCountX2 / 2;

            table.EndCode = _reader.ReadUInt16ArrayBigEndian(segCount);
            table.ReservedPad = _reader.ReadUInt16BigEndian();
            table.StartCode = _reader.ReadUInt16ArrayBigEndian(segCount);
            table.IdDelta = _reader.ReadInt16ArrayBigEndian(segCount);
            table.IdRangeOffset = _reader.ReadUInt16ArrayBigEndian(segCount);

            int bytesRead = 14 + segCount * 8 + 2;
            int glyphArrayBytes = (ushort)table.Length - bytesRead;
            int glyphCount = glyphArrayBytes / 2;

            table.GlyphIdArray = _reader.ReadUInt16ArrayBigEndian(glyphCount);

            return table;
        }
    }
}
