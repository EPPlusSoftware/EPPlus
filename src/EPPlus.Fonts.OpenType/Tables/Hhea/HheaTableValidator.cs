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
using System;

namespace EPPlus.Fonts.OpenType.Tables.Hhea
{
    public class HheaTableValidator : ITableValidator<HheaTable>
    {
        public string TableName
        {
            get { return TableNames.Hhea; }
        }

        public Type TableType => typeof(HheaTable);

        TableValidationResult ITableValidator.Validate(FontTableBase table, FontValidationContext context)
            => Validate((HheaTable)table, context);


        public TableValidationResult Validate(HheaTable table, FontValidationContext context)
        {
            var result = new TableValidationResult();
            result.TableName = TableName;


            // Rule 1: Version check with tolerant strategy
            if (table.majorVersion == HheaTableConstants.ExpectedMajorVersion &&
                table.minorVersion == HheaTableConstants.ExpectedMinorVersion)
            {
                // Expected version (e.g., 1.0) - OK
            }
            else if (table.majorVersion == 0 && table.minorVersion == 0)
            {
                // Older TrueType fonts sometimes have 0.0 - tolerate but warn
                result.AddMessage(FontValidationSeverity.Warning,
                    string.Format("Version is {0}.{1}. Expected {2}.{3}, but 0.0 is tolerated for older fonts.",
                        table.majorVersion,
                        table.minorVersion,
                        HheaTableConstants.ExpectedMajorVersion,
                        HheaTableConstants.ExpectedMinorVersion));
            }
            else
            {
                // Any other version is considered invalid
                result.AddMessage(FontValidationSeverity.Error,
                    string.Format("Invalid version: {0}.{1}. Expected {2}.{3}.",
                        table.majorVersion,
                        table.minorVersion,
                        HheaTableConstants.ExpectedMajorVersion,
                        HheaTableConstants.ExpectedMinorVersion));
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


            if (table.advanceWidthMax > context.Font.HeadTable.UnitsPerEm * 2)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"advanceWidthMax ({table.advanceWidthMax}) seems unusually large compared to HeadTable.UnitsPerEm ({context.Font.HeadTable.UnitsPerEm}).");
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
