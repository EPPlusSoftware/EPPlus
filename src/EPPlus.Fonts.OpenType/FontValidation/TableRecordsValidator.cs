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
namespace EPPlus.Fonts.OpenType.FontValidation
{
    internal class TableRecordsValidator
    {
        public TableValidationResult Validate(OpenTypeFont font, FontValidationContext context)
        {
            var result = new TableValidationResult();
            result.TableName = "TableRecords";

            var records = font.TableRecords;
            if (records == null || records.Count == 0)
            {
                result.AddMessage(FontValidationSeverity.Error, "No table records found.");
                return result;
            }

            // Total font file length from the reader
            long fileLength = font.FileLength;
            if (fileLength <= 0)
            {
                result.AddMessage(FontValidationSeverity.Error, "Font file length could not be determined or is zero.");
                // Still continue to report per-record issues if any
            }

            foreach (var kvp in records)
            {
                string tag = kvp.Key;
                TableRecord record = kvp.Value;

                // Rule 1: Tag must be 4 chars
                if (tag == null || tag.Length != 4)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        string.Format("Invalid tag '{0}'. Expected 4 characters.", tag));
                }

                // Rule 2: Offset and Length must be > 0
                if (record.Offset == 0)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        string.Format("Offset for tag '{0}' is 0.", tag));
                }
                if (record.Length == 0)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        string.Format("Length for tag '{0}' is 0.", tag));
                }

                // Rule 3: Offset must be aligned to 4-byte boundary (sfnt requirement)
                if ((record.Offset & 0x3u) != 0u)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        string.Format("Offset for tag '{0}' ({1}) is not 4-byte aligned.", tag, record.Offset));
                }

                // Rule 4: Offset must be within file bounds
                if (fileLength > 0)
                {
                    long offset = (long)record.Offset;
                    long length = (long)record.Length;

                    if (offset < 0 || offset >= fileLength)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            string.Format("Offset for tag '{0}' ({1}) is outside font file (length {2}).", tag, record.Offset, fileLength));
                    }

                    // Rule 5: Offset + Length must not exceed file length (overflow safe)
                    long end = offset + length;
                    if (end < 0 || end > fileLength)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            string.Format("Table '{0}' exceeds file bounds: offset {1} + length {2} = {3}, file length {4}.",
                                tag, record.Offset, record.Length, end, fileLength));
                    }
                }

                // Rule 6: Checksum should not be zero (soft warning for now)
                if (record.Checksum == 0u)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        string.Format("Checksum for tag '{0}' is zero.", tag));
                }
            }

            return result;
        }
    }
}
