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

namespace EPPlus.Fonts.OpenType.Tables.Hhea
{
    public class HheaTableValidator : ITableValidator<HheaTable>
    {
        public string TableName
        {
            get { return TableNames.Hhea; }
        }

        public TableValidationResult Validate(HheaTable table, FontValidationContext context)
        {
            var result = new TableValidationResult();
            result.TableName = TableName;

            // Rule 1: Version check
            if (table.majorVersion != 1 || table.minorVersion != 0)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    string.Format("Invalid version: {0}.{1}. Expected 1.0.", table.majorVersion, table.minorVersion));
            }

            // Rule 2: Ascender and Descender should not be zero
            if (table.ascender == 0)
            {
                result.AddMessage(FontValidationSeverity.Warning, "ascender is 0, which is unusual.");
            }
            if (table.descender == 0)
            {
                result.AddMessage(FontValidationSeverity.Warning, "descender is 0, which is unusual.");
            }

            // Rule 3: lineGap can be negative but usually >= 0
            if (table.lineGap < 0)
            {
                result.AddMessage(FontValidationSeverity.Information,
                    string.Format("lineGap is negative ({0}). Some platforms treat negative as zero.", table.lineGap));
            }

            // Rule 4: advanceWidthMax should be > 0
            if (table.advanceWidthMax == 0)
            {
                result.AddMessage(FontValidationSeverity.Error, "advanceWidthMax is 0, which is invalid.");
            }

            // Rule 5: metricDataFormat must be 0
            if (table.metricDataFormat != 0)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    string.Format("metricDataFormat is {0}. Expected 0.", table.metricDataFormat));
            }

            // Rule 6: numberOfHMetrics should be > 0
            if (table.numberOfHMetrics == 0)
            {
                result.AddMessage(FontValidationSeverity.Error, "numberOfHMetrics is 0, which is invalid.");
            }

            // Optional: caretSlopeRise and caretSlopeRun checks
            if (table.caretSlopeRise == 0 && table.caretSlopeRun == 0)
            {
                result.AddMessage(FontValidationSeverity.Information,
                    "caretSlopeRise and caretSlopeRun are both 0. Expected rise=1 for vertical caret.");
            }

            return result;
        }
    }
}
