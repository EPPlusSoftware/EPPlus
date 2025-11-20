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
using EPPlus.Fonts.OpenType.Tables.Glyph;
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

        public List<uint> Offsets { get; set; } = new List<uint>();
        public HeadTable.IndexToLocFormats IndexToLocFormat { get; set; }

        internal override void Clear()
        {
            Offsets.Clear();
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
            // Verify that the number of offsets matches numGlyphs + 1
            if (Offsets == null || Offsets.Count == 0)
                throw new InvalidOperationException("Offsets list cannot be null or empty.");

            // Retrieve numGlyphs from the Maxp table via the font (or inject the value)
            // Here we assume that LocaTable has a reference or that you pass it in during construction
            int expectedCount = _maxpTable.numGlyphs + 1; // or inject the value
            if (Offsets.Count != expectedCount)
                throw new InvalidOperationException($"Offsets count ({Offsets.Count}) does not match numGlyphs + 1 ({expectedCount}).");

            // Verify that offsets are sorted and not negative
            for (int i = 1; i < Offsets.Count; i++)
            {
                if (Offsets[i] < Offsets[i - 1])
                    throw new InvalidOperationException("Offsets must be in ascending order.");
            }

            // Serialization based on IndexToLocFormat
            if (IndexToLocFormat == HeadTable.IndexToLocFormats.Offset16)
            {
                foreach (var offset in Offsets)
                {
                    if (offset > 0x1FFFF) // 131072 bytes is the max for 16-bit format (since offset/2)
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

        public LocaTable CreateSubset(GlyfTable glyfTable, HeadTable.IndexToLocFormats indexToLocFormat)
        {
            if (glyfTable == null || glyfTable.Glyphs == null || glyfTable.Glyphs.Count == 0)
                throw new ArgumentNullException(nameof(glyfTable));

            var offsets = new List<uint>(glyfTable.Glyphs.Count + 1);
            uint currentOffset = 0;
            offsets.Add(currentOffset);

            foreach (var glyph in glyfTable.Glyphs)
            {
                int size = glyph.GetSize();
                currentOffset += (uint)size;
                offsets.Add(currentOffset);
            }

            // Create new LocaTable
            var newLocaTable = new LocaTable(_maxpTable)
            {
                Offsets = offsets,
                IndexToLocFormat = indexToLocFormat
            };

            return newLocaTable;
        }
    }
}
