/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/15/2025         EPPlus Software AB           Initial implementation
  01/19/2026         EPPlus Software AB           Added Single Adjustment support (GPOS Type 1)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.TextShaping.Contextual;
using EPPlus.Fonts.OpenType.TextShaping.Kerning;
using EPPlus.Fonts.OpenType.TextShaping.Ligatures;
using EPPlus.Fonts.OpenType.TextShaping.Positioning;
using EPPlus.Fonts.OpenType.TextShaping.Substitutions;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping
{
    public class TextShaper : ITextShaper
    {
        private readonly OpenTypeFont _font;
        private readonly KerningProvider _kerningProvider;
        private readonly LigatureProcessor _ligatureProcessor;
        private readonly MarkToBaseProvider _markToBaseProvider;
        private readonly SingleAdjustmentProvider _singleAdjustmentProvider;
        private readonly SingleSubstitutionProcessor _singleSubstitutionProcessor;
        private readonly ChainingContextualProcessor _chainingContextualProcessor;

        private const ushort DEFAULT_UNITS_PER_EM = 1000;

        public ushort UnitsPerEm
        {
            get
            {
                if (_font?.HeadTable?.UnitsPerEm == null || _font.HeadTable.UnitsPerEm == 0)
                {
                    return DEFAULT_UNITS_PER_EM;
                }
                return _font.HeadTable.UnitsPerEm;
            }
        }

        public TextShaper(OpenTypeFont font)
        {
            _font = font ?? throw new ArgumentNullException(nameof(font));
            _kerningProvider = new KerningProvider(font);
            _ligatureProcessor = new LigatureProcessor(font);
            _markToBaseProvider = new MarkToBaseProvider(font);
            _singleAdjustmentProvider = new SingleAdjustmentProvider(font);
            _singleSubstitutionProcessor = new SingleSubstitutionProcessor(font);
            _chainingContextualProcessor = new ChainingContextualProcessor(font, _singleSubstitutionProcessor, _ligatureProcessor);
        }

        #region Single-line Shaping

        /// <summary>
        /// Shape text using default options (ligatures + kerning).
        /// </summary>
        /// <param name="text">Text to shape</param>
        /// <returns>Shaped text with positioned glyphs</returns>
        public ShapedText Shape(string text)
        {
            return Shape(text, ShapingOptions.Default);
        }

        /// <summary>
        /// Shape text with specified options.
        /// Note: Newline characters (\n, \r, \r\n) are treated as regular characters.
        /// For multi-line text, use ShapeLines() method instead.
        /// </summary>
        /// <param name="text">Text to shape</param>
        /// <param name="options">Shaping options</param>
        /// <returns>Shaped text with positioned glyphs</returns>
        public ShapedText Shape(string text, ShapingOptions options)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new ShapedText
                {
                    OriginalText = text ?? string.Empty,
                    Glyphs = new ShapedGlyph[0]
                };
            }

            if (options == null)
            {
                options = ShapingOptions.Default;
            }

            // Phase 1: Map characters to glyphs
            var glyphs = MapToGlyphs(text);

            // Phase 2: Apply GSUB substitutions (if enabled)
            if (options.ApplySubstitutions && _font.GsubTable != null)
            {
                glyphs = ApplyGsubSubstitutions(glyphs, options);
            }

            // Phase 3: Apply GPOS positioning (if enabled)
            if (options.ApplyPositioning)
            {
                ApplyPositioning(glyphs, options);
            }

            // Phase 4: Build result
            return new ShapedText
            {
                OriginalText = text,
                Glyphs = glyphs.ToArray()
            };
        }

        public double[] ExtractCharWidths(string text, float fontSize, ShapingOptions options)
        {
            var charWidths = new double[text.Length];

            if (string.IsNullOrEmpty(text))
            {
                return charWidths;
            }

            // Shape once - entire text
            var shaped = Shape(text, options);
            double scaleFactor = fontSize / UnitsPerEm;

            // Extract widths - using ClusterIndex to map glyphs to characters
            foreach (var glyph in shaped.Glyphs)
            {
                int charIndex = glyph.ClusterIndex;

                if (charIndex >= 0 && charIndex < text.Length)
                {
                    charWidths[charIndex] += glyph.XAdvance * scaleFactor;
                }
            }

            return charWidths;
        }

        /// <summary>
        /// Extracts character widths into a pre-allocated target array to avoid new allocations.
        /// Writes widths for the first text.Length positions; caller must ensure targetArray.Length >= text.Length.
        /// </summary>
        /// <param name="text">The text to measure</param>
        /// <param name="fontSize">Font size in points</param>
        /// <param name="options">Shaping options</param>
        /// <param name="targetArray">Pre-allocated array to write widths into (must be large enough)</param>
        public void ExtractCharWidths(string text, float fontSize, ShapingOptions options, double[] targetArray)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            if (targetArray == null || targetArray.Length < text.Length)
            {
                throw new ArgumentException($"Target array must be at least as large as text length ({text.Length})", nameof(targetArray));
            }

            // Clear only the portion we will use (safer than full Array.Clear for large buffers)
            Array.Clear(targetArray, 0, text.Length);

            var shaped = Shape(text, options ?? ShapingOptions.Default);
            double scaleFactor = fontSize / UnitsPerEm;

            foreach (var glyph in shaped.Glyphs)
            {
                int charIndex = glyph.ClusterIndex;
                if (charIndex >= 0 && charIndex < text.Length)
                {
                    targetArray[charIndex] += glyph.XAdvance * scaleFactor;
                }
            }
        }



        #endregion

        #region Phase 1: Character to Glyph Mapping

        /// <summary>
        /// Maps characters to glyphs using the cmap table.
        /// </summary>
        private List<ShapedGlyph> MapToGlyphs(string text)
        {
            var glyphs = new List<ShapedGlyph>(text.Length);
            var cmapTable = _font.CmapTable;
            var hmtxTable = _font.HmtxTable;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                // Map character to glyph ID
                int glyphId = cmapTable.MapCharToGlyph(c);

                // Handle missing glyphs (use .notdef)
                if (glyphId < 0)
                {
                    glyphId = 0; // .notdef
                }

                // Get advance width from hmtx
                int advanceWidth = hmtxTable.GetAdvanceWidth((ushort)glyphId);

                glyphs.Add(new ShapedGlyph
                {
                    GlyphId = (ushort)glyphId,
                    XAdvance = advanceWidth,
                    YAdvance = 0,
                    XOffset = 0,
                    YOffset = 0,
                    ClusterIndex = i,
                    CharCount = 1
                });
            }

            return glyphs;
        }

        #endregion

        #region Phase 2: GSUB Substitutions

        /// <summary>
        /// Applies GSUB substitutions (ligatures, contextual alternates, etc.).
        /// </summary>
        private List<ShapedGlyph> ApplyGsubSubstitutions(List<ShapedGlyph> glyphs, ShapingOptions options)
        {
            // Phase 1: Single Substitution (Type 1) - applies first
            // Examples: small caps (smcp), oldstyle figures (onum), tabular figures (tnum)
            if (options.GsubFeatures != null && options.GsubFeatures.Count > 0)
            {
                glyphs = _singleSubstitutionProcessor.ApplySubstitutions(glyphs, options.GsubFeatures);
            }

            // Phase 2: Chaining Contextual Substitution (Type 6) for ligatures
            // This handles context-sensitive ligatures (e.g., ffi in Roboto)
            // Must come BEFORE simple ligatures to handle contextual cases first
            if (options.GsubFeatures != null && options.GsubFeatures.Contains("liga"))
            {
                glyphs = _chainingContextualProcessor.ApplyContextualSubstitutions(glyphs, "liga");
            }

            // Phase 3: Simple Ligatures (Type 4) - applies after contextual ligatures
            // This catches any remaining non-contextual ligatures
            if (options.GsubFeatures != null && options.GsubFeatures.Contains("liga"))
            {
                glyphs = _ligatureProcessor.ApplyLigatures(glyphs);
            }

            return glyphs;
        }

        #endregion

        #region Phase 3: Positioning

        /// <summary>
        /// Applies positioning adjustments (kerning, mark positioning, etc.).
        /// Order matters: Single adjustments → Kerning → Mark positioning
        /// </summary>
        private void ApplyPositioning(List<ShapedGlyph> glyphs, ShapingOptions options)
        {
            // Early return if positioning is disabled
            if (!options.ApplyPositioning)
            {
                return;
            }

            // Determine if we should apply all features or only specific ones
            bool applyAllFeatures = options.GposFeatures == null || options.GposFeatures.Count == 0;

            // Phase 1: Single Adjustment (GPOS Type 1)
            // ALWAYS applied when positioning is enabled - fundamental positioning
            ApplySingleAdjustment(glyphs, options);

            // Phase 2: Kerning (GPOS Type 2 / kern table)
            // Applied when: all features enabled OR "kern" is explicitly requested
            if (applyAllFeatures || (options.GposFeatures != null && options.GposFeatures.Contains("kern")))
            {
                ApplyKerning(glyphs);
            }

            // Phase 3: Mark-to-Base positioning (GPOS Type 4)
            // ALWAYS applied when positioning is enabled - critical for diacritics
            _markToBaseProvider.ApplyMarkPositioning(glyphs);
        }

        /// <summary>
        /// Applies single glyph adjustments from GPOS Lookup Type 1.
        /// Only applies adjustments from the specified features.
        /// </summary>
        private void ApplySingleAdjustment(List<ShapedGlyph> glyphs, ShapingOptions options)
        {
            // Determine which features to use
            List<string> features = options.GposFeatures ?? new List<string>();

            for (int i = 0; i < glyphs.Count; i++)
            {
                ushort glyphId = glyphs[i].GlyphId;

                if (_singleAdjustmentProvider.TryGetAdjustment(glyphId, features, out var valueRecord))
                {
                    var glyph = glyphs[i];

                    if (valueRecord.XPlacement != 0)
                        glyph.XOffset += valueRecord.XPlacement;
                    if (valueRecord.YPlacement != 0)
                        glyph.YOffset += valueRecord.YPlacement;
                    if (valueRecord.XAdvance != 0)
                        glyph.XAdvance += valueRecord.XAdvance;
                    if (valueRecord.YAdvance != 0)
                        glyph.YAdvance += valueRecord.YAdvance;

                    glyphs[i] = glyph;
                }
            }
        }

        /// <summary>
        /// Applies kerning adjustments to glyph pairs.
        /// Uses KerningProvider which handles both GPOS and legacy kern table.
        /// </summary>
        private void ApplyKerning(List<ShapedGlyph> glyphs)
        {
            for (int i = 1; i < glyphs.Count; i++)
            {
                ushort leftGlyph = glyphs[i - 1].GlyphId;
                ushort rightGlyph = glyphs[i].GlyphId;

                // Get kerning value (handles GPOS + kern table + caching)
                short kernValue = _kerningProvider.GetKerning(leftGlyph, rightGlyph);

                if (kernValue != 0)
                {
                    // Apply kerning to the left glyph's advance
                    var glyph = glyphs[i - 1];
                    glyph.XAdvance += kernValue;
                    glyphs[i - 1] = glyph;
                }
            }
        }

        #endregion

        #region Utilities

        /// <summary>
        /// Measures the width of text in font units.
        /// </summary>
        public int MeasureText(string text, ShapingOptions options = null)
        {
            var shaped = Shape(text, options);
            return shaped.TotalAdvanceWidth;
        }

        /// <summary>
        /// Measures the width of text in PDF points.
        /// </summary>
        public float MeasureTextInPoints(string text, float fontSize, ShapingOptions options = null)
        {
            var shaped = Shape(text, options);
            float unitsPerEm = _font.HeadTable.UnitsPerEm;
            return shaped.GetWidthInPoints(fontSize, unitsPerEm);
        }

        /// <summary>
        /// Measures the width of text in pixels.
        /// </summary>
        public float MeasureTextInPixels(string text, float fontSize, float dpi, ShapingOptions options = null)
        {
            var shaped = Shape(text, options);
            float unitsPerEm = _font.HeadTable.UnitsPerEm;
            return shaped.GetWidthInPixels(fontSize, dpi, unitsPerEm);
        }

        #endregion

        #region Multi-line Support

        /// <summary>
        /// Shape multi-line text (handles \n, \r, \r\n).
        /// Returns one ShapedText per line.
        /// </summary>
        public ShapedText[] ShapeLines(string text, ShapingOptions options = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new ShapedText[0];
            }

            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var result = new ShapedText[lines.Length];

            for (int i = 0; i < lines.Length; i++)
            {
                result[i] = Shape(lines[i], options);
            }

            return result;
        }

        /// <summary>
        /// Measure multi-line text and return bounding box.
        /// </summary>
        public MultiLineMetrics MeasureLines(string text, float fontSize, ShapingOptions options = null)
        {
            var shapedLines = ShapeLines(text, options);
            float unitsPerEm = _font.HeadTable.UnitsPerEm;

            float maxWidth = 0;
            foreach (var line in shapedLines)
            {
                float lineWidth = line.GetWidthInPoints(fontSize, unitsPerEm);
                maxWidth = Math.Max(maxWidth, lineWidth);
            }

            float lineHeight = GetLineHeightInPoints(fontSize);
            float fontHeight = GetFontHeightInPoints(fontSize);
            float totalHeight = shapedLines.Length * lineHeight;

            return new MultiLineMetrics
            {
                Width = maxWidth,
                Height = totalHeight,
                FontHeight = fontHeight,
                LineCount = shapedLines.Length,
                LineHeight = lineHeight
            };
        }

        /// <summary>
        /// Get line height (ascent + descent + line gap) in points.
        /// </summary>
        public float GetLineHeightInPoints(float fontSize)
        {
            var hhea = _font.HheaTable;
            float unitsPerEm = _font.HeadTable.UnitsPerEm;

            // ascent is positive, descender is negative
            int lineHeightUnits = hhea.ascender - hhea.descender + hhea.lineGap;

            return (lineHeightUnits / unitsPerEm) * fontSize;
        }

        /// <summary>
        /// Get font height (ascent + descent only, no line gap) in points.
        /// </summary>
        public float GetFontHeightInPoints(float fontSize)
        {
            var hhea = _font.HheaTable;
            float unitsPerEm = _font.HeadTable.UnitsPerEm;

            // ascent is positive, descender is negative
            int fontHeightUnits = hhea.ascender - hhea.descender;

            return (fontHeightUnits / unitsPerEm) * fontSize;
        }

        #endregion

        /// <summary>
        /// Gets single line spacing (baseline-to-baseline distance).
        /// Uses typo metrics if USE_TYPO_METRICS flag is set, otherwise uses Win metrics.
        /// </summary>
        public double GetLineHeightInPoints(double fontSize)
        {
            if (_font.Os2Table.UseTypoMetrics)
            {
                // Modern fonts: use typo metrics
                var typoAscent = _font.Os2Table.sTypoAscender;
                var typoDescent = _font.Os2Table.sTypoDescender;
                var typoLineGap = _font.Os2Table.sTypoLineGap;
                double em = _font.HeadTable.UnitsPerEm;
                double lineHeight = typoAscent - typoDescent + typoLineGap;
                return (lineHeight / em) * fontSize;
            }
            else
            {
                // Legacy fonts: use Win metrics (same as font height)
                return GetFontHeightInPoints(fontSize);
            }
        }

        // ✅ HAR REDAN
        public double GetFontHeightInPoints(double fontSize)
        {
            // Total font height (ascent + descent)
            var ascent = _font.Os2Table.usWinAscent;
            var descent = _font.Os2Table.usWinDescent;
            var em = _font.HeadTable.UnitsPerEm;

            return (ascent + descent) * (fontSize / em);
        }

        // ❌ SAKNAS - LÄGG TILL DENNA!
        public double GetBaseLineInPoints(double fontSize)
        {
            // Distance from top of box to baseline
            var ascent = _font.Os2Table.UseTypoMetrics
                ? (double)_font.Os2Table.sTypoAscender
                : (double)_font.Os2Table.usWinAscent;

            var em = _font.HeadTable.UnitsPerEm;
            return ascent * (fontSize / em);
        }

        // BONUS - Kan vara användbart
        public double GetDescentInPoints(double fontSize)
        {
            var descent = _font.Os2Table.UseTypoMetrics
                ? (double)Math.Abs(_font.Os2Table.sTypoDescender)  // Descent är negativ
                : _font.Os2Table.usWinDescent;

            var em = _font.HeadTable.UnitsPerEm;
            return descent * (fontSize / em);
        }
    }
}