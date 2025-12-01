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
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Maxp;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Loca
{
    /// <summary>
    /// The indexToLoc table stores the offsets to the locations of the glyphs in the font, relative to the beginning of the glyphData table
    /// https://docs.microsoft.com/en-us/typography/opentype/spec/loca
    /// </summary>
    public class LocaTable : FontTableBase
    {
        public LocaTable(MaxpTable maxpTable)
        {
            _maxpTable = maxpTable;
        }

        private readonly MaxpTable _maxpTable;

        public override string Name => TableNames.Loca;

        public override bool IsEssentialTable => true;

        public List<uint> Offsets { get; set; } = new List<uint>();
        public HeadTable.IndexToLocFormats IndexToLocFormat { get; set; }

        internal override void Clear()
        {
            Offsets.Clear();
        }



        internal int GetGlyphCountSafe()
        {
            // According to OpenType spec:
            // loca table contains numGlyphs + 1 entries (last entry = end of last glyph)
            // So glyph count = Offsets.Count - 1 if Offsets is valid.
            // If Offsets is null or empty, return 0 as a safe fallback.

            if (Offsets == null || Offsets.Count < 2)
            {
                return 0; // No valid glyph offsets
            }

            // Normally, Offsets.Count should equal _maxpTable.numGlyphs + 1
            // But we trust Offsets.Count for safety and subtract 1.
            return Offsets.Count - 1;
        }



        internal static LocaTable CreateSubset(List<uint> offsets, HeadTable.IndexToLocFormats indexToLocFormat, MaxpTable maxpTable)
        {
            var newLocaTable = new LocaTable(maxpTable)
            {
                Offsets = offsets,
                IndexToLocFormat = indexToLocFormat
            };
            return newLocaTable;
        }


        internal override void SerializeInternal(FontsBinaryWriter writer)
        {
            // Kontrollera att antalet offsets matchar numGlyphs + 1
            if (Offsets == null || Offsets.Count == 0)
                throw new InvalidOperationException("Offsets list cannot be null or empty.");

            // Hämta numGlyphs från Maxp-tabellen via fonten (eller injicera värdet)
            // Här antar vi att LocaTable har en referens eller att du skickar in det vid konstruktion
            int expectedCount = _maxpTable.numGlyphs + 1; // eller injicera värdet
            if (Offsets.Count != expectedCount)
                throw new InvalidOperationException($"Offsets count ({Offsets.Count}) does not match numGlyphs + 1 ({expectedCount}).");

            // Kontrollera att offsets är sorterade och inte negativa
            for (int i = 1; i < Offsets.Count; i++)
            {
                if (Offsets[i] < Offsets[i - 1])
                    throw new InvalidOperationException("Offsets must be in ascending order.");
            }

            // Serialisering baserat på IndexToLocFormat
            if (IndexToLocFormat == HeadTable.IndexToLocFormats.Offset16)
            {
                foreach (var offset in Offsets)
                {
                    if (offset > 0x1FFFF) // 131072 bytes är max för 16-bit format (eftersom offset/2)
                        throw new InvalidOperationException($"Offset {offset} exceeds maximum allowed for Offset16 format.");

                    ushort shortOffset = (ushort)(offset / 2);
                    writer.WriteUInt16BigEndian(shortOffset);
                }
            }
            else if (IndexToLocFormat == HeadTable.IndexToLocFormats.Offset32)
            {
                foreach (var offset in Offsets)
                {
                    writer.WriteUInt32BigEndian(offset);
                }
            }
            else
            {
                throw new InvalidOperationException($"Unsupported IndexToLocFormat: {IndexToLocFormat}");
            }
        }
    }
}
