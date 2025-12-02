using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Head;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Loca
{

    internal class LocaTableValidator : TableValidatorBase<LocaTable>
    {
        public override Type TableType => typeof(LocaTable);
        public override string TableName => "loca";

        public override TableValidationResult Validate(LocaTable loca, FontValidationContext context)
        {
            var result = new TableValidationResult();
            result.TableName = TableName;
            var font = context.Font;

            // 1. Check that offsets count matches numGlyphs + 1
            int expectedCount = font.MaxpTable.numGlyphs + 1;
            if (loca.Offsets.Count != expectedCount)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Loca: Expected {expectedCount} offsets (numGlyphs + 1), found {loca.Offsets.Count}.");
            }

            // 2. Validate IndexToLocFormat
            if (loca.IndexToLocFormat != HeadTable.IndexToLocFormats.Offset16 &&
                loca.IndexToLocFormat != HeadTable.IndexToLocFormats.Offset32)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Loca: Invalid IndexToLocFormat {loca.IndexToLocFormat}. Must be Offset16 or Offset32.");
            }

            var glyfTableLength = font.GlyfTable.GetLength();

            // 3. Validate offsets ascending and glyph lengths
            for (int i = 0; i < loca.Offsets.Count - 1; i++)
            {
                uint current = loca.Offsets[i];
                uint next = loca.Offsets[i + 1];

                if (next < current)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Loca: Offset[{i}] ({current}) > Offset[{i + 1}] ({next}). Offsets must be ascending.");
                }

                uint glyphLength = next - current;
                if (glyphLength > glyfTableLength)
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        $"Loca: Glyph {i} length {glyphLength} exceeds glyf table size ({glyfTableLength}).");
                }
            }

            // 4. Validate all offsets within glyf table bounds
            foreach (var offset in loca.Offsets)
            {
                if (offset > glyfTableLength)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Loca: Offset {offset} exceeds glyf table length ({glyfTableLength}).");
                }

                // Additional check for Offset16 format
                if (loca.IndexToLocFormat == HeadTable.IndexToLocFormats.Offset16 && offset > 0x1FFFF)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Loca: Offset {offset} exceeds maximum allowed for Offset16 format (131072 bytes).");
                }
            }

            // 5. Optional: Warn if glyf table is empty but offsets > 0
            if (glyfTableLength == 0 && loca.Offsets.Any(o => o > 0))
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    "Loca: Glyf table is empty but offsets contain non-zero values.");
            }


            // 6. Check that IndexToLocFormat in loca matches the head table
            if (loca.IndexToLocFormat != font.HeadTable.IndexToLocFormat)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Loca: IndexToLocFormat mismatch with head table. Expected {font.HeadTable.IndexToLocFormat}, found {loca.IndexToLocFormat}.");
            }

            // 7. Check that the last offset ≈ the end of the glyf table
            uint lastOffset = loca.Offsets.Last();
            if (lastOffset != glyfTableLength)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"Loca: Last offset ({lastOffset}) does not match glyf table length ({glyfTableLength}).");
            }


            return result;
        }
    }

}
