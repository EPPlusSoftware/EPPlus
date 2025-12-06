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
        // TA BORT HELT: private readonly MaxpTable _maxpTable;

        public LocaTable()
        {
        }

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
            if (Offsets == null || Offsets.Count < 2)
                return 0;

            return Offsets.Count - 1;
        }

        // ENDA factoryn – ingen Maxp-referens längre
        internal static LocaTable CreateSubset(List<uint> offsets, HeadTable.IndexToLocFormats indexToLocFormat)
        {
            return new LocaTable
            {
                Offsets = new List<uint>(offsets), // defensiv kopia
                IndexToLocFormat = indexToLocFormat
            };
        }

        internal override void SerializeInternal(FontsBinaryWriter writer, FontSerializationContext context)
        {
            if (Offsets == null || Offsets.Count == 0)
                throw new InvalidOperationException("Offsets list cannot be null or empty.");

            // NYTT: Hämta numGlyphs från fonten (via context) istället för intern referens
            int expectedCount = context.Font?.MaxpTable?.numGlyphs + 1 ?? Offsets.Count;

            if (Offsets.Count != expectedCount)
            {
                // VIKTIGT: För subset är detta OK ibland (t.ex. under byggande)
                // Men vi loggar bara – kastar inte i subset-läge
                if (context.IsSubsetInProgress != true)
                {
                    throw new InvalidOperationException(
                        $"Offsets count ({Offsets.Count}) does not match expected numGlyphs + 1 ({expectedCount}).");
                }
            }

            // Resten oförändrad – perfekt som den är
            for (int i = 1; i < Offsets.Count; i++)
            {
                if (Offsets[i] < Offsets[i - 1])
                    throw new InvalidOperationException("Offsets must be in ascending order.");
            }

            if (IndexToLocFormat == HeadTable.IndexToLocFormats.Offset16)
            {
                foreach (var offset in Offsets)
                {
                    if (offset > 0x1FFFF)
                        throw new InvalidOperationException($"Offset {offset} exceeds maximum for Offset16 format.");

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
