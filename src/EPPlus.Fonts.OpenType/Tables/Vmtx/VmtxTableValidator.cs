/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/19/2026         EPPlus Software AB           vmtx table implementation (vertical text support)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontValidation;
using System;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Vmtx
{
    /// <summary>
    /// Validates the 'vmtx' (Vertical Metrics) table.
    /// This is an optional table, only present in fonts with vertical metrics (primarily CJK).
    /// Analogous to HmtxTableValidator.
    /// </summary>
    internal class VmtxTableValidator : TableValidatorBase<VmtxTable>
    {
        public override Type TableType => typeof(VmtxTable);
        public override string TableName => TableNames.Vmtx;

        public override TableValidationResult Validate(VmtxTable table, FontValidationContext context)
        {
            var result = new TableValidationResult { TableName = TableName, LogLevel = base.LogLevel };

            var vhea = context.Font.VheaTable;
            var maxp = context.Font.MaxpTable;

            // 1. vhea is required to interpret vmtx
            if (vhea == null)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    "vmtx table is present but vhea table is missing. Cannot validate vmtx without vhea.");
                return result;
            }

            // 2. VMetrics count must match vhea.NumberOfVMetrics
            int vMetricsCount = table.VMetrics?.Count ?? 0;
            if (vMetricsCount != vhea.NumberOfVMetrics)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"vmtx.VMetrics count ({vMetricsCount}) does not match vhea.NumberOfVMetrics ({vhea.NumberOfVMetrics}).");
            }

            // 3. Total glyph coverage must equal maxp.numGlyphs
            if (maxp != null)
            {
                int tsbCount = table.TopSideBearings?.Count ?? 0;
                int totalCoverage = vMetricsCount + tsbCount;
                if (totalCoverage != maxp.numGlyphs)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"vmtx total glyph coverage ({totalCoverage}) does not match maxp.numGlyphs ({maxp.numGlyphs}). " +
                        $"VMetrics ({vMetricsCount}) + TopSideBearings ({tsbCount}) should equal {maxp.numGlyphs}.");
                }
            }

            // 4. All advance heights must be > 0
            if (table.VMetrics != null)
            {
                for (int i = 0; i < table.VMetrics.Count; i++)
                {
                    if (table.VMetrics[i].AdvanceHeight == 0)
                    {
                        result.AddMessage(FontValidationSeverity.Warning,
                            $"vmtx.VMetrics[{i}].AdvanceHeight is 0, which is unusual.");
                    }
                }
            }

            // 5. AdvanceHeightMax in vhea should match the actual max in vmtx
            if (table.VMetrics != null && table.VMetrics.Count > 0)
            {
                ushort maxAdvanceHeight = table.VMetrics.Max(m => m.AdvanceHeight);
                if (vhea.AdvanceHeightMax < maxAdvanceHeight)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"vhea.AdvanceHeightMax ({vhea.AdvanceHeightMax}) is less than the actual max advanceHeight in vmtx ({maxAdvanceHeight}).");
                }
            }

            return result;
        }
    }
}