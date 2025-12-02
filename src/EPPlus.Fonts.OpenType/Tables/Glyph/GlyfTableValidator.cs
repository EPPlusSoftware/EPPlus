using EPPlus.Fonts.OpenType.FontValidation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
                if (glyph == null)
                {
                    // Check if loca offsets indicate an empty glyph
                    if (loca != null && i + 1 < loca.Offsets.Count)
                    {
                        if (loca.Offsets[i] == loca.Offsets[i + 1])
                        {
                            result.AddMessage(FontValidationSeverity.Information,
                                $"Glyph {i} is empty (null) and loca offsets confirm zero length.");
                        }
                        else
                        {
                            result.AddMessage(FontValidationSeverity.Warning,
                                $"Glyph {i} is null but loca offsets indicate data exists.");
                        }
                    }
                    else
                    {
                        result.AddMessage(FontValidationSeverity.Warning,
                            $"Glyph {i} is null. Cannot confirm via loca offsets.");
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
                    if (glyph.SimpleData != null || glyph.CompositeData != null)
                        result.AddMessage(FontValidationSeverity.Warning, $"Glyph {i} has contour count 0 but contains data.");
                }

                // Validate glyph size vs loca offsets or alignment
                if (loca != null && i + 1 < loca.Offsets.Count)
                {
                    int actualSize = glyph.GetSize();
                    var diff = loca.Offsets[i + 1] - loca.Offsets[i];

                    if (actualSize > diff)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            $"Glyph {i} size ({actualSize}) exceeds loca offset range ({diff}).");
                    }
                    else if (actualSize < diff)
                    {
                        result.AddMessage(FontValidationSeverity.Information,
                            $"Glyph {i} size ({actualSize}) is smaller than loca offset range ({diff}). Padding detected.");
                    }
                }
                else
                {
                    int size = glyph.GetSize();
                    if (size % 4 != 0)
                    {
                        result.AddMessage(FontValidationSeverity.Warning,
                            $"Glyph {i} size ({size}) is not 4-byte aligned.");
                    }
                }
            }

            return result;
        }





        private void ValidateSimpleGlyph(SimpleGlyph simple, int glyphIndex, TableValidationResult result)
        {
            // Check that EndPtsOfContours exists
            if (simple.EndPtsOfContours == null || simple.EndPtsOfContours.Length == 0)
            {
                result.AddMessage(FontValidationSeverity.Error, $"Glyph {glyphIndex} has no EndPtsOfContours.");
            }
            else
            {
                int lastEndPt = simple.EndPtsOfContours.Last();
                int pointCount = simple.Points?.Count ?? 0;

                if (pointCount == 0)
                {
                    // Points are not loaded, skip consistency check but continue with other validations
                    result.AddMessage(FontValidationSeverity.Warning,
                        $"Glyph {glyphIndex} has EndPtsOfContours but no points loaded. Skipping point consistency check.");
                }
                else
                {
                    // Validate that last end point matches point count - 1
                    if (lastEndPt != pointCount - 1)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            $"Glyph {glyphIndex} EndPtsOfContours last value ({lastEndPt}) does not match point count - 1 ({pointCount - 1}).");
                    }

                    // Validate bounding box against actual points
                    short minX = simple.Points.Min(p => p.X);
                    short maxX = simple.Points.Max(p => p.X);
                    short minY = simple.Points.Min(p => p.Y);
                    short maxY = simple.Points.Max(p => p.Y);

                    if (minX < -32768 || maxX > 32767 || minY < -32768 || maxY > 32767)
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            $"Glyph {glyphIndex} has point coordinates out of valid range (-32768 to 32767).");
                    }
                }
            }

            // Validate flags count matches point count (if points are loaded)
            int pointCountForFlags = simple.Points?.Count ?? 0;
            if (pointCountForFlags > 0 && simple.Flags.Count != pointCountForFlags)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Glyph {glyphIndex} has {simple.Flags.Count} flags but {pointCountForFlags} points.");
            }

            // Validate instruction length does not exceed 65535
            if (simple.Instructions.Length > 65535)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Glyph {glyphIndex} has instruction length {simple.Instructions.Length}, exceeds 65535.");
            }

            // Validate coordinate byte arrays exist if points are present
            if ((simple.XBytes == null || simple.YBytes == null) && (simple.Points?.Count ?? 0) > 0)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Glyph {glyphIndex} has points but missing XBytes or YBytes.");
            }
        }


        private void ValidateCompositeGlyph(CompositeGlyph composite, int glyphIndex, TableValidationResult result, int glyphCount)
        {
            // Check that there is at least one component
            if (composite.Components == null || composite.Components.Count == 0)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    $"Glyph {glyphIndex} is composite but has no components.");
                return;
            }

            // Validate each component
            for (int c = 0; c < composite.Components.Count; c++)
            {
                var comp = composite.Components[c];

                // Check glyph index reference
                if (comp.GlyphIndex >= glyphCount)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Glyph {glyphIndex} component {c} references invalid glyph index {comp.GlyphIndex}.");
                }

                // Validate flag combinations
                bool hasScale = (comp.Flags & CompositeGlyphFlags.WE_HAVE_A_SCALE) != 0;
                bool hasXYScale = (comp.Flags & CompositeGlyphFlags.WE_HAVE_AN_X_AND_Y_SCALE) != 0;
                bool hasTwoByTwo = (comp.Flags & CompositeGlyphFlags.WE_HAVE_A_TWO_BY_TWO) != 0;

                int scaleFlagsCount = (hasScale ? 1 : 0) + (hasXYScale ? 1 : 0) + (hasTwoByTwo ? 1 : 0);
                if (scaleFlagsCount > 1)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Glyph {glyphIndex} component {c} has conflicting scale flags.");
                }

                // Validate transformation values (F2Dot14 range: -4.0 to +4.0)
                if (hasScale && !IsF2Dot14Valid(comp.Scale))
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Glyph {glyphIndex} component {c} has invalid Scale value {comp.Scale}.");
                }

                if (!IsF2Dot14Valid(comp.XScale) || !IsF2Dot14Valid(comp.YScale))
                {
                    if (IsF2Dot14Safe(comp.XScale) && IsF2Dot14Safe(comp.YScale))
                    {
                        result.AddMessage(FontValidationSeverity.Warning,
                            $"Glyph {glyphIndex} component {c} has XScale/YScale outside spec range but within short range: XScale={comp.XScale}, YScale={comp.YScale}.");
                    }
                    else
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            $"Glyph {glyphIndex} component {c} has invalid XScale/YScale values: XScale={comp.XScale}, YScale={comp.YScale}.");
                    }
                }

                if (hasTwoByTwo && (!IsF2Dot14Valid(comp.XScale) || !IsF2Dot14Valid(comp.YScale) ||
                                    !IsF2Dot14Valid(comp.Scale01) || !IsF2Dot14Valid(comp.Scale10)))
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Glyph {glyphIndex} component {c} has invalid 2x2 transformation values.");
                }
            }

            // Validate instructions if WE_HAVE_INSTRUCTIONS flag is set
            var lastComp = composite.Components.Last();
            if ((lastComp.Flags & CompositeGlyphFlags.WE_HAVE_INSTRUCTIONS) != 0)
            {
                if (composite.Instructions == null || composite.Instructions.Length == 0)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Glyph {glyphIndex} has WE_HAVE_INSTRUCTIONS flag but no instructions.");
                }
                else if (composite.Instructions.Length > 65535)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        $"Glyph {glyphIndex} instructions length exceeds 65535.");
                }
            }
        }

        // Helper method to validate F2Dot14 range
        private bool IsF2Dot14Valid(short value)
        {
            // F2Dot14 format: signed 16-bit, 2 integer bits, 14 fractional bits
            // Valid range: -4.0 to +4.0 → raw value between -16384 and +16384
            return value >= -16384 && value <= 16384;
        }


        private bool IsF2Dot14Safe(short value)
        {
            // Allow full short range but warn if outside spec
            return value >= short.MinValue && value <= short.MaxValue;
        }
    }
}
