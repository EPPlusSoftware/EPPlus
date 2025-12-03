using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType
{

    internal class OpenTypeFontSerializer
    {
        private readonly OpenTypeFont _font;

        public OpenTypeFontSerializer(OpenTypeFont font)
        {
            _font = font ?? throw new ArgumentNullException(nameof(font));
        }

        public byte[] Serialize()
        {
            using var stream = new MemoryStream();
            using var writer = new FontsBinaryWriter(stream);

            var sortedRecords = _font.TableRecords
                .OrderBy(r => r.Key)
                .Select(r => r.Value)
                .ToList();

            int numTables = sortedRecords.Count;

            // 1. Write sfnt header
            WriteSfntHeader(writer, numTables);

            // 2. Write table directory
            foreach (var record in sortedRecords)
            {
                WriteTableRecord(writer, record);
            }

            // 3. Write tables (use preprocessed padded bytes if available)
            foreach (var record in sortedRecords)
            {
                var tag = record.Tag.Value;

                byte[] tableBytes;
                if (_font.PreprocessedPaddedTables != null &&
                    _font.PreprocessedPaddedTables.TryGetValue(tag, out var cachedBytes))
                {
                    tableBytes = cachedBytes; // Use preprocessed padded bytes
                }
                else
                {
                    // Fallback: get raw data and pad
                    tableBytes = _font.GetTableData(tag);
                    int rawLen = tableBytes.Length;
                    int paddedLen = (rawLen + 3) & ~3;
                    if (paddedLen > rawLen)
                    {
                        Array.Resize(ref tableBytes, paddedLen);
                    }
                }

                writer.Write(tableBytes);
            }

            return stream.ToArray();
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
