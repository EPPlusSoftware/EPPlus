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

namespace EPPlus.Fonts.OpenType
{

    internal class OpenTypeFontSerializer
    {
        private readonly OpenTypeFont _font;

        public OpenTypeFontSerializer(OpenTypeFont font)
        {
            _font = font ?? throw new ArgumentNullException(nameof(font));
        }

        // In OpenTypeFontSerializer class

        public byte[] Serialize()
        {
            using (var stream = new MemoryStream())
            using (var writer = new FontsBinaryWriter(stream))
            {
                var sortedTags = _font.TableRecords.Keys.OrderBy(k => k).ToList();
                int numTables = sortedTags.Count;

                // 1. First pass: serialize all tables to get their actual bytes
                var tableBytes = new Dictionary<string, byte[]>();
                foreach (var tag in sortedTags)
                {
                    byte[] data;
                    if (_font.PreprocessedPaddedTables != null &&
                        _font.PreprocessedPaddedTables.TryGetValue(tag, out var cachedBytes))
                    {
                        data = cachedBytes;
                    }
                    else
                    {
                        data = _font.GetTableData(tag);
                        // Pad to 4-byte boundary
                        int rawLen = data.Length;
                        int paddedLen = (rawLen + 3) & ~3;
                        if (paddedLen > rawLen)
                        {
                            Array.Resize(ref data, paddedLen);
                        }
                    }
                    tableBytes[tag] = data;
                }

                // 2. Calculate correct offsets
                // Header = 12 bytes, each table record = 16 bytes
                uint currentOffset = (uint)(12 + numTables * 16);

                var newRecords = new List<TableRecord>();
                foreach (var tag in sortedTags)
                {
                    var data = tableBytes[tag];
                    var originalRecord = _font.TableRecords[tag];

                    var newRecord = new TableRecord
                    {
                        Tag = new Tag(tag),
                        Checksum = originalRecord.Checksum, // Keep original checksum for now
                        Offset = currentOffset,
                        Length = originalRecord.Length
                    };
                    newRecords.Add(newRecord);

                    // Move to next table (already padded)
                    currentOffset += (uint)data.Length;
                }

                // 3. Write sfnt header
                WriteSfntHeader(writer, numTables);

                // 4. Write table directory with CORRECT offsets
                foreach (var record in newRecords)
                {
                    WriteTableRecord(writer, record);
                }

                // 5. Write table data
                foreach (var tag in sortedTags)
                {
                    writer.Write(tableBytes[tag]);
                }

                return stream.ToArray();
            }
        }

        private void WriteSfntHeader(FontsBinaryWriter writer, int numTables)
        {
            writer.WriteUInt32BigEndian(0x00010000); // sfntVersion for TrueType
            writer.WriteUInt16BigEndian((ushort)numTables);

            int maxPower2 = 1;
            int entrySelector = 0;
            while (maxPower2 * 2 <= numTables)
            {
                maxPower2 *= 2;
                entrySelector++;
            }
            ushort searchRange = (ushort)(maxPower2 * 16);
            ushort rangeShift = (ushort)(numTables * 16 - searchRange);

            writer.WriteUInt16BigEndian(searchRange);
            writer.WriteUInt16BigEndian((ushort)entrySelector);
            writer.WriteUInt16BigEndian(rangeShift);
        }

        private void WriteTableRecord(FontsBinaryWriter writer, TableRecord record)
        {
            writer.Write(record.Tag.ToBytes()); // 4 bytes ASCII
            writer.WriteUInt32BigEndian(record.Checksum);
            writer.WriteUInt32BigEndian(record.Offset);
            writer.WriteUInt32BigEndian(record.Length);
        }
    }

}