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
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    /// <summary>
    /// This table defines the mapping of character codes to the glyph index values used in the font. It may contain more than one subtable, in order to support more than one character encoding scheme.
    /// </summary>
    public class CmapTable : FontTableBase
    {
        internal CmapTable()
        {
            EncodingRecords = new List<EncodingRecord>();
            SubTables = new List<CmapSubtableBase>();
        }

        public override string Name => TableNames.Cmap;

        public override bool IsEssentialTable => true;

        /// <summary>
        /// Table version number (0).
        /// </summary>
        public ushort Version { get; set; }

        /// <summary>
        /// Number of encoding tables that follow.
        /// </summary>
        public ushort NumTables { get; set; }

        /// <summary>
        /// The array of encoding records specifies particular encodings and the offset to the subtable for each encoding.
        /// </summary>
        public List<EncodingRecord> EncodingRecords { get; private set; }

        /// <summary>
        /// Array of Subtables
        /// </summary>
        public List<CmapSubtableBase> SubTables { get; private set; }



        internal override void SerializeInternal(FontsBinaryWriter writer, FontSerializationContext context)
        {
            // Start of cmap table
            long tableStart = writer.BaseStream.Position;

            // Write header
            writer.WriteUInt16BigEndian(Version);
            writer.WriteUInt16BigEndian((ushort)EncodingRecords.Count);

            // Reserve space for encoding records
            long encodingRecordStart = writer.BaseStream.Position;
            foreach (var _ in EncodingRecords)
            {
                writer.Write(new byte[8]); // placeholder
            }

            // Precompute offsets for unique subtables. Deduplication is keyed by the SUBTABLE
            // INSTANCE, not by SubtableOffset: for a freshly-built cmap (e.g. a subset's cmap),
            // every EncodingRecord starts out with the same placeholder SubtableOffset (0), so
            // keying on that value would wrongly alias two DIFFERENT subtables (e.g. a subset's
            // format 4 and format 12 tables) onto the same bytes the first time this ran with two
            // fresh tables sharing that placeholder. Keying by the actual object reference only
            // dedups encoding records that genuinely point at the SAME subtable (e.g. (3,1) and
            // (0,3) both referencing one shared Unicode BMP subtable), which is what this is for.
            var subtableOffsetsMap = new Dictionary<CmapSubtableBase, uint>();
            var subTableStartIndex = writer.BaseStream.Position;
            var encRecordsToSerialize = EncodingRecords.OrderBy(er => er.SubtableOffset);
            foreach (var encRecord in encRecordsToSerialize)
            {
                // Skip any explicitly marked skipped records. Format 14 (Unicode Variation
                // Sequences) subtables ARE serialized like any other format now - they used to be
                // unconditionally dropped here, which silently discarded variation-sequence data
                // from every embedded font.
                if (encRecord.IsSkipped)
                {
                    continue;
                }
                uint existingOffset;
                if (encRecord.Subtable != null && subtableOffsetsMap.TryGetValue(encRecord.Subtable, out existingOffset))
                {
                    encRecord.SubtableOffset = existingOffset;
                    continue;
                }
                var subTableBytes = encRecord.Subtable.Serialize();
                writer.Write(subTableBytes);
                if (encRecord.Subtable != null)
                    subtableOffsetsMap[encRecord.Subtable] = (uint)subTableStartIndex;
                encRecord.SubtableOffset = (uint)subTableStartIndex;
                subTableStartIndex += subTableBytes.Length;

            }

            // Go back and write encoding records with correct offsets
            long currentPos = writer.BaseStream.Position;
            writer.BaseStream.Seek(encodingRecordStart, SeekOrigin.Begin);

            for (int i = 0; i < EncodingRecords.Count; i++)
            {
                var record = EncodingRecords[i];
                writer.WriteUInt16BigEndian((ushort)record.PlatformId);
                writer.WriteUInt16BigEndian(record.EncodingId);
                writer.WriteUInt32BigEndian(record.SubtableOffset);
            }

            // Return to end of stream
            writer.BaseStream.Seek(currentPos, SeekOrigin.Begin);
        }


        public int MapCharToGlyph(char ch)
        {
            int codePoint = ch; // Unicode value
            foreach (var subtable in SubTables)
            {
                int glyphId = subtable.MapCodePointToGlyph(codePoint);
                if (glyphId >= 0)
                    return glyphId;
            }
            return -1; // Not found
        }


        public int GetMinCharCode()
        {
            int minCode = int.MaxValue;

            // Defensive: handle empty/none
            if (SubTables == null || SubTables.Count == 0)
                return 0;

            for (int i = 0; i < SubTables.Count; i++)
            {
                CmapSubtableBase sub = SubTables[i];
                if (sub == null) continue;

                var mappings = sub.GetGlyphMappings();
                if (mappings == null || mappings.CharCodeToGlyphIndex == null) continue;

                // Iterate all char-code → glyph-index pairs
                foreach (KeyValuePair<uint, ushort> kvp in mappings.CharCodeToGlyphIndex)
                {
                    // glyphIndex is ushort, so it's always >= 0; we only need the char code
                    uint code = kvp.Key;
                    if (code < (uint)minCode)
                    {
                        minCode = (int)code;
                    }
                }
            }

            // If no mappings found, return 0
            return (minCode == int.MaxValue) ? 0 : minCode;
        }

        public int GetMaxCharCode()
        {
            int maxCode = int.MinValue;

            if (SubTables == null || SubTables.Count == 0)
                return 0;

            for (int i = 0; i < SubTables.Count; i++)
            {
                CmapSubtableBase sub = SubTables[i];
                if (sub == null) continue;

                var mappings = sub.GetGlyphMappings();
                if (mappings == null || mappings.CharCodeToGlyphIndex == null) continue;

                foreach (KeyValuePair<uint, ushort> kvp in mappings.CharCodeToGlyphIndex)
                {
                    uint code = kvp.Key;
                    if (code > (uint)maxCode)
                    {
                        maxCode = (int)code;
                    }
                }
            }

            return (maxCode == int.MinValue) ? 0 : maxCode;
        }


        public bool ContainsChar(ushort charCode)
        {
            foreach (var subtable in SubTables)
            {
                if (subtable.TryGetGlyphId(charCode, out _))
                {
                    return true;
                }
            }
            return false;
        }


        public bool TryGetGlyphId(uint codePoint, out ushort glyphId)
        {
            glyphId = 0;

            var preferred = GetPreferredSubtable();
            if (preferred != null && preferred.TryGetGlyphId(codePoint, out glyphId) && glyphId != 0)
            {
                return true;
            }

            // Fallback: loopa alla
            foreach (var subtable in SubTables)
            {
                if (subtable.TryGetGlyphId(codePoint, out glyphId) && glyphId != 0)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Looks up a Unicode Variation Sequence - a (base character, variation selector) pair -
        /// against the font's cmap format 14 subtable (Unicode Variation Sequences, see the
        /// OpenType spec's "Format 14" section). Returns true only if the sequence is actually
        /// registered in the font:
        ///   - a "non-default" entry supplies an explicit override glyph for the base character, or
        ///   - a "default" entry means the sequence is registered but carries no glyph of its own -
        ///     the base character's ordinary glyph (as <see cref="TryGetGlyphId(uint, out ushort)"/>
        ///     would return) should be used.
        /// Returns false when there is no format 14 subtable at all, the variation selector isn't
        /// registered in it, or the selector is registered but this particular base character is not
        /// listed under it. In every false case the pair is not a known variation sequence, and the
        /// caller should fall back to treating the base character on its own.
        /// </summary>
        public bool TryGetGlyphId(uint baseCodePoint, uint variationSelector, out ushort glyphId)
        {
            glyphId = 0;

            CmapSubtable14 subtable14 = null;
            foreach (var subtable in SubTables)
            {
                if (subtable.Format == 14)
                {
                    subtable14 = subtable as CmapSubtable14;
                    break;
                }
            }
            if (subtable14 == null)
                return false;

            foreach (var selector in subtable14.VariationSelectors)
            {
                if (selector.VarSelector != variationSelector)
                    continue;

                if (selector.NonDefaultUvsTable != null)
                {
                    foreach (var mapping in selector.NonDefaultUvsTable.Mappings)
                    {
                        if (mapping.UnicodeValue == baseCodePoint)
                        {
                            glyphId = mapping.GlyphId;
                            return true;
                        }
                    }
                }

                if (selector.DefaultUvsTable != null)
                {
                    foreach (var range in selector.DefaultUvsTable.Ranges)
                    {
                        uint rangeEnd = range.StartUnicodeValue + (uint)range.AdditionalCount;
                        if (baseCodePoint >= range.StartUnicodeValue && baseCodePoint <= rangeEnd)
                        {
                            // A "default" entry carries no glyph of its own - it just confirms the
                            // sequence is registered, so fall back to the base character's ordinary glyph.
                            return TryGetGlyphId(baseCodePoint, out glyphId);
                        }
                    }
                }

                // The selector itself is registered in this font, but this base character is not
                // listed under it in either table - not a known sequence.
                return false;
            }

            // No entry at all for this variation selector.
            return false;
        }

        public CmapSubtableBase GetPreferredSubtable()
        {

            // Prioritetsordning: Format 12 > Format 4 > Format 6 > Format 0
            var preferredFormats = new ushort[] { 12, 4, 6, 0 };

            foreach (var format in preferredFormats)
            {
                for (int i = 0; i < EncodingRecords.Count; i++)
                {
                    var record = EncodingRecords[i];
                    if (record.PlatformId == Platforms.Windows && record.EncodingId == 1)
                    {
                        var subtable = EncodingRecords[i].Subtable;
                        if (subtable != null && subtable.Format == format)
                        {
                            return subtable;
                        }
                    }
                }
            }

            return null;

        }

        internal override void Clear()
        {
            NumTables = 0;
            EncodingRecords.Clear();
            SubTables.Clear();
        }

        public uint GetUnicodeCodePoint(ushort glyphId)
        {
            foreach (var subtable in SubTables)
            {
                var mappings = subtable.GetGlyphMappings();
                if (mappings == null) continue;

                foreach (var kvp in mappings.CharCodeToGlyphIndex)
                {
                    if (kvp.Value == glyphId) return kvp.Key;
                }
            }
            return 0;
        }

        public ushort GetGlyphId(char ch)
        {
            TryGetGlyphId(ch, out ushort gid);
            return gid;
        }
    }
}