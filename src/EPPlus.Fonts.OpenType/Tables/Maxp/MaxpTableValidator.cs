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

namespace EPPlus.Fonts.OpenType.Tables.Maxp
{
    public class MaxpTableValidator : ITableValidator<MaxpTable>
    {
        public string TableName
        {
            get { return TableNames.Maxp; }
        }

        public TableValidationResult Validate(MaxpTable table, FontValidationContext context)
        {
            var result = new TableValidationResult();
            result.TableName = TableName;

            // Rule 1: numGlyphs must be > 0
            if (table.numGlyphs <= 0)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    string.Format("numGlyphs is {0}. Must be greater than 0.", table.numGlyphs));
            }

            // Rule 2: version must be 0x00005000 (0.5) or 0x00010000 (1.0)
            if (table.version.RawValue != 0x00005000 && table.version.RawValue != 0x00010000)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    string.Format("Unexpected version: 0x{0:X}. Expected 0x00005000 or 0x00010000.", table.version.RawValue));
            }

            // Rule 3: Version-specific checks
            if (table.version.RawValue == 0x00010000)
            {
                // maxZones should be 1 or 2 (per spec text, most fonts use 2).
                if (table.maxZones != 1 && table.maxZones != 2)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        string.Format("maxZones is {0}. Expected 1 or 2.", table.maxZones));
                }

                // (Optional sanity) Ensure extended fields are not obviously inconsistent.
                // Note: We do NOT enforce composite >= non-composite; that is not a spec requirement.
                // All fields are ushort, so they are non-negative by type.

                // If you want a soft check on extremes, you could add informational notes
                // for extremely large values, but avoid strict comparisons or false positives.
            }
            else if (table.version.RawValue == 0x00005000)
            {
                // Version 0.5 only defines numGlyphs; extended fields should be zero.
                if (table.maxPoints != 0 || table.maxContours != 0 ||
                    table.maxCompositePoints != 0 || table.maxCompositeContours != 0 ||
                    table.maxZones != 0 || table.maxTwilightPoints != 0 ||
                    table.maxStorage != 0 || table.maxFunctionDefs != 0 ||
                    table.maxInstructionDefs != 0 || table.maxStackElements != 0 ||
                    table.maxSizeOfInstructions != 0 || table.maxComponentElements != 0 ||
                    table.maxComponentDepth != 0)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        "Version 0.5 should not define extended fields (they should be zero).");
                }
            }

            return result;
        }
    }
}
