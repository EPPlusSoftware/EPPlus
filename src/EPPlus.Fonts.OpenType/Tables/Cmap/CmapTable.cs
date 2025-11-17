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

        public CmapTable CreateSubset(IEnumerable<char> usedChars, Dictionary<ushort, ushort> idMapping)
        {
            var newCmap = new CmapTable { Version = 0 };

            // Build mappings: charCode -> new glyph ID
            List<CharGlyphMapping> mappings = new List<CharGlyphMapping>();
            foreach (char ch in usedChars)
            {
                int oldGlyphId = this.MapCharToGlyph(ch);
                if (oldGlyphId >= 0 && idMapping.ContainsKey((ushort)oldGlyphId))
                {
                    ushort newGlyphId = idMapping[(ushort)oldGlyphId];
                    mappings.Add(new CharGlyphMapping((uint)ch, newGlyphId));
                }
            }

            mappings.Sort(delegate (CharGlyphMapping a, CharGlyphMapping b)
            {
                return a.CharCode.CompareTo(b.CharCode);
            });

            CmapSubtableBase newSubtable;
            CmapSubtableBase preferred = GetPreferredSubtable();
            if (preferred == null)
                throw new InvalidOperationException("No suitable cmap subtable found.");

            switch (preferred.Format)
            {
                case 12:
                    CmapSubtable12 sub12 = new CmapSubtable12();
                    sub12.BuildFromMappings(mappings);
                    newSubtable = sub12;
                    break;

                case 4:
                    CmapSubtable4 sub4 = new CmapSubtable4();
                    sub4.BuildFromMappings(mappings);
                    newSubtable = sub4;
                    break;

                case 6:
                    CmapSubtable6 sub6 = new CmapSubtable6();
                    sub6.BuildFromMappings(mappings);
                    newSubtable = sub6;
                    break;

                case 0:
                    CmapSubtable0 sub0 = new CmapSubtable0();
                    sub0.BuildFromMappings(mappings);
                    newSubtable = sub0;
                    break;

                default:
                    throw new NotSupportedException("Cmap format " + preferred.Format + " not supported for subset.");
            }

            EncodingRecord record = new EncodingRecord(Platforms.Windows, 1, 0);
            record.Subtable = newSubtable;

            newCmap.EncodingRecords.Add(record);
            newCmap.SubTables.Add(newSubtable);
            newCmap.NumTables = (ushort)newCmap.EncodingRecords.Count;

            return newCmap;
        }
    }
}
