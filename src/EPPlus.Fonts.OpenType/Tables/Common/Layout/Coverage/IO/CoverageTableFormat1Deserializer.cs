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
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage.IO
{
    internal class CoverageTableFormat1Deserializer
    {
        private readonly FontsBinaryReader _reader;

        public CoverageTableFormat1Deserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public CoverageTableFormat1 Deserialize(long startIndex)
        {
            _reader.BaseStream.Seek(startIndex, SeekOrigin.Begin);

            // Read Format (already known to be 1, but read to advance position)
            ushort format = _reader.ReadUInt16BigEndian();

            CoverageTableFormat1 table = new CoverageTableFormat1 { CoverageFormat = format };

            // USHORT GlyphCount
            table.GlyphCount = _reader.ReadUInt16BigEndian();

            // USHORT[] GlyphArray
            table.GlyphArray = new ushort[table.GlyphCount];

            for (int i = 0; i < table.GlyphCount; i++)
            {
                table.GlyphArray[i] = _reader.ReadUInt16BigEndian();
            }

            return table;
        }
    }
}
