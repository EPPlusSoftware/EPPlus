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
using EPPlus.Fonts.OpenType.FontValidation;

namespace EPPlus.Fonts.OpenType.Tables.Name
{
    public class NameTableValidator : ITableValidator<NameTable>
    {
        public string TableName
        {
            get { return TableNames.Name; }
        }

        public TableValidationResult Validate(NameTable table, FontValidationContext context)
        {
            var result = new TableValidationResult();
            result.TableName = TableName;

            // Rule 1: Table must contain records
            if (table.NameRecords == null || table.NameRecords.Length == 0)
            {
                result.AddMessage(FontValidationSeverity.Error, "No name records found in name table.");
                return result;
            }

            bool hasFamilyName = false;
            bool hasSubfamilyName = false;

            for (int i = 0; i < table.NameRecords.Length; i++)
            {
                var record = table.NameRecords[i];

                // Required NameIDs
                if (record.nameId == 1) // Font Family
                {
                    hasFamilyName = true;
                    if (string.IsNullOrEmpty(record.Name))
                    {
                        result.AddMessage(FontValidationSeverity.Error, "Font Family name (NameID=1) is empty.");
                    }
                }
                else if (record.nameId == 2) // Font Subfamily
                {
                    hasSubfamilyName = true;
                    if (string.IsNullOrEmpty(record.Name))
                    {
                        result.AddMessage(FontValidationSeverity.Error, "Font Subfamily name (NameID=2) is empty.");
                    }
                }

                // Optional: Encoding consistency for Windows platform
                if (record.platformId == 3 && record.encodingId != 1 && record.encodingId != 10)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        string.Format("Unexpected EncodingID {0} for Windows platform (expected 1 or 10).", record.encodingId));
                }
            }

            if (!hasFamilyName)
            {
                result.AddMessage(FontValidationSeverity.Error, "Missing required Font Family name (NameID=1).");
            }
            if (!hasSubfamilyName)
            {
                result.AddMessage(FontValidationSeverity.Error, "Missing required Font Subfamily name (NameID=2).");
            }

            return result;
        }
    }
}
