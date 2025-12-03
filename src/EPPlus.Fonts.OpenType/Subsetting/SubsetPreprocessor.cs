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

using EPPlus.Fonts.OpenType.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    internal class SubsetPreprocessor
    {
        public void PreprocessSubset(OpenTypeFont font)
        {
            // 1) Sort table records by tag (spec requires directory sorted alphabetically)
            var sortedRecords = font.TableRecords
                .OrderBy(r => r.Key)
                .Select(r => r.Value)
                .ToList();

            int numTables = sortedRecords.Count;
            uint offset = (uint)(12 + numTables * 16); // sfnt header (12) + table directory (16 * numTables)

            // Ensure we have a cache for padded bytes used by serializer
            font.PreprocessedPaddedTables?.Clear();

            // 2) For checksum total, head.checkSumAdjustment must be 0 initially
            font.HeadTable.ChecksumAdjustment = 0;

            // 3) Serialize each table, pad to 4 bytes, compute checksum via ChecksumCalculator
            foreach (var record in sortedRecords)
            {
                string tag = record.Tag.Value;

                byte[] rawBytes = font.GetTableData(tag);   // unpadded
                int rawLen = rawBytes?.Length ?? 0;

                // Pad to 4-byte boundary for checksum and layout
                int paddedLen = (rawLen + 3) & ~3;
                if (paddedLen > rawLen)
                {
                    Array.Resize(ref rawBytes, paddedLen); // zero-padding
                }

                // Compute checksum using the shared calculator (handles head zeroing)
                uint checksum = ChecksumCalculator.CalculateTableChecksum(rawBytes, tag);

                // Update TableRecord
                record.Offset = offset;
                record.Length = (uint)(rawLen);      // length must be UNPADDED per spec
                record.Checksum = checksum;

                // Cache padded bytes for the serializer to write verbatim
                font.PreprocessedPaddedTables[tag] = rawBytes;

                // Advance offset by padded length
                offset += (uint)paddedLen;
            }

            // 4) Compute the font-wide checksum with head.checkSumAdjustment == 0
            uint totalSum = ComputeFontChecksum(font, sortedRecords, font.PreprocessedPaddedTables);
            uint adjustment = 0xB1B0AFBA - totalSum;

            // 5) Update head.checkSumAdjustment, re-serialize head table and recompute its checksum
            font.HeadTable.ChecksumAdjustment = adjustment;

            // Re-serialize HEAD (un-padded), then pad for checksum and cache
            byte[] headRaw = font.GetTableData("head");       // unpadded AFTER adjustment was set
            int headRawLen = headRaw.Length;
            int headPaddedLen = (headRawLen + 3) & ~3;
            if (headPaddedLen > headRawLen)
            {
                Array.Resize(ref headRaw, headPaddedLen);
            }

            // Recompute head checksum using shared calculator (which zeroes bytes 8–11)
            uint headChecksum = ChecksumCalculator.CalculateTableChecksum(headRaw, "head");
            var headRecord = sortedRecords.First(r => r.Tag.Value == "head");
            headRecord.Checksum = headChecksum;

            // Update cache with the final padded HEAD bytes (with adjustment included)
            font.PreprocessedPaddedTables["head"] = headRaw;
        }

        private uint ComputeFontChecksum(OpenTypeFont font,
                                         List<TableRecord> records,
                                         Dictionary<string, byte[]> paddedTables)
        {
            // Header
            byte[] header = BuildSfntHeader(records.Count);
            uint sum = ChecksumCalculator.CalculateTableChecksum(header, ""); // tag unused here

            // Table directory (concatenate directory entries)
            foreach (var record in records)
            {
                byte[] dirEntry = BuildTableRecordBytes(record); // 16 bytes
                sum += ChecksumCalculator.CalculateTableChecksum(dirEntry, "");
            }

            // Tables (already padded)
            foreach (var kvp in paddedTables)
            {
                sum += ChecksumCalculator.CalculateTableChecksum(kvp.Value, kvp.Key);
            }

            return sum;
        }

        private byte[] BuildSfntHeader(int numTables)
        {
            using var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);

            writer.WriteUInt32BigEndian(0x00010000); // TrueType sfntVersion
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

            return ms.ToArray();
        }

        private byte[] BuildTableRecordBytes(TableRecord record)
        {
            using var ms = new MemoryStream();
            using var writer = new FontsBinaryWriter(ms);

            writer.Write(record.Tag.ToBytes());          // 4 bytes ASCII
            writer.WriteUInt32BigEndian(record.Checksum);
            writer.WriteUInt32BigEndian(record.Offset);
            writer.WriteUInt32BigEndian(record.Length);

            return ms.ToArray();
        }
    }
}
