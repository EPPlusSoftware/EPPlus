using EPPlus.Fonts.OpenType.FontValidation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Hmtx
{

    public class HmtxTableValidator : ITableValidator<HmtxTable>
    {
        public Type TableType => typeof(HmtxTable);
        public string TableName => "hmtx";

        public TableValidationResult Validate(FontTableBase table, FontValidationContext context)
        {
            return Validate((HmtxTable)table, context);
        }

        public TableValidationResult Validate(HmtxTable hmtx, FontValidationContext context)
        {
            var result = new TableValidationResult();
            result.TableName = TableName;
            var font = context.Font;

            int numGlyphs = font.MaxpTable.numGlyphs;
            int numHMetrics = font.HheaTable.numberOfHMetrics;

            // 1. Validate hMetrics count
            if (hmtx.hMetrics.Count != numHMetrics)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Hmtx: Expected {numHMetrics} hMetrics entries (from hhea), found {hmtx.hMetrics.Count}.");
            }

            // 2. Validate total coverage
            int totalEntries = hmtx.hMetrics.Count + hmtx.leftSideBearings.Count;
            if (totalEntries < numGlyphs)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Hmtx: Total entries ({totalEntries}) do not cover all glyphs ({numGlyphs}).");
            }

            // 3. Validate advance widths
            for (int i = 0; i < hmtx.hMetrics.Count; i++)
            {
                ushort aw = hmtx.hMetrics[i].advanceWidth;
                if (aw == 0)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        $"Hmtx: advanceWidth for glyph {i} is 0 (may be intentional, but check).");
                }
            }

            // 4. Validate LSB values
            foreach (var metric in hmtx.hMetrics)
            {
                if (metric.lsb < -32768 || metric.lsb > 32767)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Hmtx: LSB value {metric.lsb} out of valid range (-32768 to 32767).");
                }
            }

            foreach (var lsb in hmtx.leftSideBearings)
            {
                if (lsb < -32768 || lsb > 32767)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Hmtx: LSB value {lsb} out of valid range (-32768 to 32767).");
                }
            }

            return result;
        }
    }

}
