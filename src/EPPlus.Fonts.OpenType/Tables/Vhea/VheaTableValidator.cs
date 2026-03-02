/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/18/2026         EPPlus Software AB           vhea table implementation (vertical text support)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontValidation;
using System;

namespace EPPlus.Fonts.OpenType.Tables.Vhea
{
    /// <summary>
    /// Validates the 'vhea' (Vertical Header) table.
    /// This is an optional table, only present in fonts with vertical metrics (primarily CJK).
    /// </summary>
    internal class VheaTableValidator : TableValidatorBase<VheaTable>
    {
        public override Type TableType => typeof(VheaTable);
        public override string TableName => TableNames.Vhea;

        public override TableValidationResult Validate(VheaTable table, FontValidationContext context)
        {
            var result = new TableValidationResult { TableName = TableName, LogLevel = base.LogLevel };

            // 1. Version check: must be 1.0 (0x00010000) or 1.1 (0x00011000)
            if (table.Version != 0x00010000 && table.Version != 0x00011000)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"vhea version is 0x{table.Version:X8}. Expected 0x00010000 (v1.0) or 0x00011000 (v1.1).");
            }

            // 2. Ascent and Descent should not be zero individually
            if (table.Ascent == 0)
            {
                result.AddMessage(FontValidationSeverity.Warning, "Ascent is 0, which is unusual.");
            }
            if (table.Descent == 0)
            {
                result.AddMessage(FontValidationSeverity.Warning, "Descent is 0, which is unusual.");
            }

            // 3. Sanity-check against unitsPerEm
            var head = context.Font.HeadTable;
            if (head != null)
            {
                int totalHeight = table.Ascent - table.Descent;
                if (totalHeight > head.UnitsPerEm * 2)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        $"vhea total height (Ascent - Descent = {totalHeight}) is more than 2x unitsPerEm ({head.UnitsPerEm}). This is unusual.");
                }
            }

            // 3. LineGap can be negative but usually >= 0
            if (table.LineGap < 0)
            {
                result.AddMessage(FontValidationSeverity.Information,
                    $"LineGap is negative ({table.LineGap}). Some platforms treat negative as zero.");
            }

            // 4. AdvanceHeightMax must be positive
            if (table.AdvanceHeightMax == 0)
            {
                result.AddMessage(FontValidationSeverity.Error, "AdvanceHeightMax is 0, which is invalid.");
            }

            if (head != null && table.AdvanceHeightMax > head.UnitsPerEm * 2)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"AdvanceHeightMax ({table.AdvanceHeightMax}) seems unusually large compared to HeadTable.UnitsPerEm ({head.UnitsPerEm}).");
            }

            // 5. MetricDataFormat must be 0 (only defined value per spec)
            if (table.MetricDataFormat != 0)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"vhea.MetricDataFormat is {table.MetricDataFormat}. Only 0 is defined by the spec.");
            }

            // 6. NumberOfVMetrics must be > 0 and consistent with maxp.numGlyphs
            if (table.NumberOfVMetrics == 0)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    "vhea.NumberOfVMetrics is 0. The vmtx table cannot be parsed without at least one entry.");
            }
            else
            {
                var maxp = context.Font.MaxpTable;
                if (maxp != null && table.NumberOfVMetrics > maxp.numGlyphs)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"vhea.NumberOfVMetrics ({table.NumberOfVMetrics}) exceeds maxp.numGlyphs ({maxp.numGlyphs}). Font is likely corrupt.");
                }
            }

            // 7. Reserved fields should be 0
            if (table.Reserved1 != 0 || table.Reserved2 != 0 ||
                table.Reserved3 != 0 || table.Reserved4 != 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    "One or more vhea reserved fields are non-zero. These should be 0 per spec.");
            }

            return result;
        }
    }
}