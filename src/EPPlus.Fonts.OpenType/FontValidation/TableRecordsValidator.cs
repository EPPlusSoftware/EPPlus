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

namespace EPPlus.Fonts.OpenType.FontValidation
{
    internal class TableRecordsValidator
    {
        public TableValidationResult Validate(OpenTypeFont font, FontValidationContext context, FontValidationSeverity logLevel)
        {
            var result = new TableValidationResult { TableName = "TableRecords", LogLevel = logLevel };

            var records = font.TableRecords;
            if (records == null || records.Count == 0)
            {
                result.AddMessage(FontValidationSeverity.Error, "No table records found.");
                return result;
            }

            long fileLength = font.FileLength;
            if (fileLength <= 0)
            {
                result.AddMessage(FontValidationSeverity.Error, "Font file length could not be determined or is zero.");
            }

            foreach (var kvp in records)
            {
                string tag = kvp.Key;
                TableRecord record = kvp.Value;

                // Rule 1: Tag must be 4 chars
                if (tag == null || tag.Length != 4)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Invalid tag '{tag}'. Expected 4 characters.");
                }

                // Rule 2: Offset and Length must be > 0
                if (record.Offset == 0)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Offset for tag '{tag}' is 0.");
                }
                if (record.Length == 0)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Length for tag '{tag}' is 0.");
                }

                // Rule 3: Offset must be aligned to 4-byte boundary
                if ((record.Offset & 0x3u) != 0u)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        $"Offset for tag '{tag}' ({record.Offset}) is not 4-byte aligned.");
                }

                // Rule 4: Offset must be within file bounds
                if (fileLength > 0)
                {
                    long offset = (long)record.Offset;
                    long length = (long)record.Length;

                    if (offset < 0 || offset >= fileLength)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            $"Offset for tag '{tag}' ({record.Offset}) is outside font file (length {fileLength}).");
                    }

                    long end = offset + length;
                    if (end < 0 || end > fileLength)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            $"Table '{tag}' exceeds file bounds: offset {record.Offset} + length {record.Length} = {end}, file length {fileLength}.");
                    }
                }

                // Rule 5: Validate checksum using ChecksumCalculator
                var tableData = font.GetTableData(tag);
                if (tableData != null && tableData.Length > 0)
                {
                    uint calculatedChecksum = ChecksumCalculator.CalculateTableChecksum(tableData, tag);
                    if (calculatedChecksum != record.Checksum)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            $"Checksum mismatch for table '{tag}': expected {record.Checksum}, got {calculatedChecksum}.");
                    }
                }
                else
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        $"Could not read data for table '{tag}' to validate checksum.");
                }
            }

            // Validate font-level checksum adjustment
            var headTable = font.HeadTable;
            if (headTable != null)
            {
                byte[] fontData = (byte[])font.RawData.Clone();
                int adjustmentOffset = (int)records["head"].Offset + 8; // checkSumAdjustment är vid offset 8
                for (int i = 0; i < 4; i++) fontData[adjustmentOffset + i] = 0;

                uint sum = ChecksumCalculator.CalculateFontChecksum(fontData);
                uint expectedAdjustment = 0xB1B0AFBA - sum;

                if (expectedAdjustment != headTable.ChecksumAdjustment)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Font checksum adjustment failed: expected {expectedAdjustment:X8}, got {headTable.ChecksumAdjustment:X8}.");
                }
            }
            else
            {
                result.AddMessage(FontValidationSeverity.Warning, "Head table missing, cannot validate font checksum adjustment.");
            }


            return result;
        }
    }
}
