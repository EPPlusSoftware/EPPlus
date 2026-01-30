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
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Glyph
{

    internal class GlyfTableValidator : TableValidatorBase<GlyfTable>
    {
        public override Type TableType => typeof(GlyfTable);
        public override string TableName => TableNames.Glyf;

        public override TableValidationResult Validate(GlyfTable table, FontValidationContext context)
        {
            var result = new TableValidationResult { TableName = TableName, LogLevel = base.LogLevel };

            // 1. Validate dependencies
            var loca = context.Font.LocaTable;
            var maxp = context.Font.MaxpTable;

            if (loca == null)
                result.AddMessage(FontValidationSeverity.Error, "Missing loca table required by glyf.");
            if (maxp == null)
                result.AddMessage(FontValidationSeverity.Error, "Missing maxp table required by glyf.");

            if (maxp != null)
            {
                if (table.Glyphs.Count < maxp.numGlyphs)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Glyph count ({table.Glyphs.Count}) is less than maxp.numGlyphs ({maxp.numGlyphs}). Font may be broken.");
                }
                else if (table.Glyphs.Count > maxp.numGlyphs)
                {
                    result.AddMessage(FontValidationSeverity.Information,
                        $"Glyph count ({table.Glyphs.Count}) is greater than maxp.numGlyphs ({maxp.numGlyphs}). This is allowed but unusual.");
                }
            }

            if (loca != null && maxp != null)
            {
                if (loca.Offsets.Count < maxp.numGlyphs + 1)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Loca offsets count ({loca.Offsets.Count}) is less than maxp.numGlyphs + 1 ({maxp.numGlyphs + 1}). Font may be broken.");
                }
                else if (loca.Offsets.Count != table.Glyphs.Count + 1)
                {
                    result.AddMessage(FontValidationSeverity.Information,
                        $"Loca offsets count ({loca.Offsets.Count}) does not match glyph count + 1 ({table.Glyphs.Count + 1}). This is allowed but unusual.");
                }
            }

            // 2. Validate glyphs
            for (int i = 0; i < table.Glyphs.Count; i++)
            {
                var glyph = table.Glyphs[i];

                // Check for empty/null glyphs using loca offsets
                if (glyph == null)
                {
                    if (loca != null && i + 1 < loca.Offsets.Count)
                    {
                        var lengthInLoca = loca.Offsets[i + 1] - loca.Offsets[i];
                        if (lengthInLoca == 0)
                        {
                            // This is a legitimate empty glyph (like 'space')
                            result.AddMessage(FontValidationSeverity.Information,
                                $"Glyph {i} is empty (null) as confirmed by zero-length entry in loca table.");
                        }
                        else
                        {
                            result.AddMessage(FontValidationSeverity.Error,
                                $"Glyph {i} is null, but loca offsets indicate it should contain {lengthInLoca} bytes of data.");
                        }
                    }
                    else
                    {
                        result.AddMessage(FontValidationSeverity.Warning,
                            $"Glyph {i} is null and cannot be verified against loca table.");
                    }
                    continue;
                }

                var header = glyph.Header;
                if (header.xMin > header.xMax || header.yMin > header.yMax)
                    result.AddMessage(FontValidationSeverity.Error, $"Glyph {i} has invalid bounding box.");

                if (header.numberOfContours > 0)
                {
                    if (glyph.SimpleData == null)
                    {
                        result.AddMessage(FontValidationSeverity.Error, $"Glyph {i} is simple but has no SimpleData.");
                    }
                    else
                    {
                        ValidateSimpleGlyph(glyph.SimpleData, i, result);
                    }
                }
                else if (header.numberOfContours < 0)
                {
                    if (glyph.CompositeData == null)
                        result.AddMessage(FontValidationSeverity.Error, $"Glyph {i} is composite but has no CompositeData.");
                    else
                        ValidateCompositeGlyph(glyph.CompositeData, i, result, table.Glyphs.Count);
                }
                else
                {
                    // numberOfContours == 0: perfectly valid for whitespace
                    if (glyph.SimpleData != null || glyph.CompositeData != null)
                        result.AddMessage(FontValidationSeverity.Warning, $"Glyph {i} has contour count 0 but contains attached data objects.");
                }

                // Validate actual serialized size vs loca table
                if (loca != null && i + 1 < loca.Offsets.Count)
                {
                    int actualSize = glyph.GetSize();
                    uint expectedSize = loca.Offsets[i + 1] - loca.Offsets[i];

                    if (actualSize > expectedSize)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            $"Glyph {i} serialized size ({actualSize}) exceeds space allocated in loca table ({expectedSize}).");
                    }
                    else if (actualSize < expectedSize)
                    {
                        // This is usually fine (padding), but good to log as info
                        result.AddMessage(FontValidationSeverity.Information,
                            $"Glyph {i} size ({actualSize}) is smaller than loca range ({expectedSize}). Padding: {expectedSize - actualSize} bytes.");
                    }
                }
            }

            return result;
        }

        private void ValidateSimpleGlyph(SimpleGlyph simple, int glyphIndex, TableValidationResult result)
        {
            if (simple.EndPtsOfContours == null || simple.EndPtsOfContours.Length == 0)
            {
                result.AddMessage(FontValidationSeverity.Error, $"Glyph {glyphIndex} has no EndPtsOfContours.");
                return;
            }

            int actualPointCount = simple.Points?.Count ?? 0;
            bool hasBinaryData = (simple.XBytes != null && simple.XBytes.Length > 0) ||
                                 (simple.YBytes != null && simple.YBytes.Length > 0);

            if (actualPointCount == 0 && !hasBinaryData)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"Glyph {glyphIndex} has EndPtsOfContours but no point list or binary coordinate data loaded.");
            }
            else if (actualPointCount > 0)
            {
                // Full consistency check for decoded points
                if (simple.EndPtsOfContours.Last() != actualPointCount - 1)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Glyph {glyphIndex} EndPtsOfContours last index ({simple.EndPtsOfContours.Last()}) " +
                        $"mismatch with point count ({actualPointCount}).");
                }

                short minX = simple.Points.Min(p => p.X);
                short maxX = simple.Points.Max(p => p.X);
                short minY = simple.Points.Min(p => p.Y);
                short maxY = simple.Points.Max(p => p.Y);

                if (minX < -32768 || maxX > 32767 || minY < -32768 || maxY > 32767)
                {
                    result.AddMessage(FontValidationSeverity.Error, $"Glyph {glyphIndex} coordinates out of 16-bit range.");
                }

                if (simple.Flags != null && simple.Flags.Count != actualPointCount)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Glyph {glyphIndex} has {simple.Flags.Count} flags but {actualPointCount} points.");
                }
            }

            if (simple.Instructions != null && simple.Instructions.Length > 65535)
            {
                result.AddMessage(FontValidationSeverity.Error, $"Glyph {glyphIndex} instruction length exceeds 65535.");
            }
        }

        private void ValidateCompositeGlyph(CompositeGlyph composite, int glyphIndex, TableValidationResult result, int glyphCount)
        {
            if (composite.Components == null || composite.Components.Count == 0)
            {
                result.AddMessage(FontValidationSeverity.Error, $"Glyph {glyphIndex} is composite but has no components.");
                return;
            }

            for (int c = 0; c < composite.Components.Count; c++)
            {
                var comp = composite.Components[c];
                if (comp.GlyphIndex >= glyphCount)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Glyph {glyphIndex} component {c} references invalid GID {comp.GlyphIndex}.");
                }

                // Flag validation
                bool hasScale = (comp.Flags & CompositeGlyphFlags.WE_HAVE_A_SCALE) != 0;
                bool hasXYScale = (comp.Flags & CompositeGlyphFlags.WE_HAVE_AN_X_AND_Y_SCALE) != 0;
                bool hasTwoByTwo = (comp.Flags & CompositeGlyphFlags.WE_HAVE_A_TWO_BY_TWO) != 0;

                if ((hasScale ? 1 : 0) + (hasXYScale ? 1 : 0) + (hasTwoByTwo ? 1 : 0) > 1)
                {
                    result.AddMessage(FontValidationSeverity.Error, $"Glyph {glyphIndex} component {c} has conflicting scale flags.");
                }
            }

            var lastComp = composite.Components.Last();
            if ((lastComp.Flags & CompositeGlyphFlags.WE_HAVE_INSTRUCTIONS) != 0)
            {
                if (composite.Instructions == null || composite.Instructions.Length == 0)
                    result.AddMessage(FontValidationSeverity.Error, $"Glyph {glyphIndex} missing instructions despite flag.");
            }
        }

        private bool IsF2Dot14Valid(short value) => value >= -16384 && value <= 16384;
        private bool IsF2Dot14Safe(short value) => value >= short.MinValue && value <= short.MaxValue;
    }
}
