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



        internal override void SerializeInternal(FontsBinaryWriter writer)
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

            // Precompute offsets for unique subtables
            var subtableOffsetsMap = new Dictionary<CmapSubtableBase, uint>();
            var subTableStartIndex = writer.BaseStream.Position;
            var encRecordsToSerialize = EncodingRecords.OrderBy(er => er.SubtableOffset);
            var usedSubtables = new Dictionary<uint, uint>();
            foreach(var encRecord in encRecordsToSerialize)
            {
                if (usedSubtables.ContainsKey(encRecord.SubtableOffset))
                {
                    encRecord.SubtableOffset = usedSubtables[encRecord.SubtableOffset];
                    continue;
                }
                var subTableBytes = encRecord.Subtable.Serialize();
                writer.Write(subTableBytes);
                usedSubtables.Add(encRecord.SubtableOffset, (uint)subTableStartIndex);
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


        internal bool TryGetGlyphId(int codePoint, out ushort glyphId)
        {
            foreach (var subtable in SubTables)
            {
                if (subtable.TryGetGlyphId(codePoint, out glyphId))
                {
                    return true;
                }
            }

            glyphId = 0;
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
    }
}
