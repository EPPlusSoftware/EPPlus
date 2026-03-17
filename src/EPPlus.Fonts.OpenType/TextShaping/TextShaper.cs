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
  02/05/2026         EPPlus Software AB           Added IFontProvider support for fallback fonts
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.TextShaping.Contextual;
using EPPlus.Fonts.OpenType.TextShaping.Kerning;
using EPPlus.Fonts.OpenType.TextShaping.Ligatures;
using EPPlus.Fonts.OpenType.TextShaping.Positioning;
using EPPlus.Fonts.OpenType.TextShaping.Substitutions;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping
{
    public partial class TextShaper : ITextShaper
    {
        private readonly OpenTypeFont _primaryFont;
        private readonly KerningProvider _kerningProvider;
        private readonly LigatureProcessor _ligatureProcessor;
        private readonly MarkToBaseProvider _markToBaseProvider;
        private readonly SingleAdjustmentProvider _singleAdjustmentProvider;
        private readonly SingleSubstitutionProcessor _singleSubstitutionProcessor;
        private readonly ChainingContextualProcessor _chainingContextualProcessor;
        private readonly IFontProvider _fontProvider;

        // Font tracking for multi-font support
        private readonly Dictionary<OpenTypeFont, byte> _fontToIdMap = new Dictionary<OpenTypeFont, byte>();
        private readonly List<OpenTypeFont> _usedFonts = new List<OpenTypeFont>();

        private const ushort DEFAULT_UNITS_PER_EM = 1000;

        public ushort UnitsPerEm
        {
            get
            {
                if (_primaryFont?.HeadTable?.UnitsPerEm == null || _primaryFont.HeadTable.UnitsPerEm == 0)
                {
                    return DEFAULT_UNITS_PER_EM;
                }
                return _primaryFont.HeadTable.UnitsPerEm;
            }
        }

        /// <summary>
        /// Creates a TextShaper with automatic emoji fallback (DefaultFontProvider).
        /// </summary>
        public TextShaper(OpenTypeFont font)
            : this(new DefaultFontProvider(font))
        {
        }

        /// <summary>
        /// Creates a TextShaper with custom font provider.
        /// </summary>
        public TextShaper(IFontProvider fontProvider)
        {
            if (fontProvider == null)
                throw new ArgumentNullException("fontProvider");
            if (fontProvider.PrimaryFont == null)
                throw new ArgumentException("Primary font cannot be null in font provider", "fontProvider");

            var gposTable = fontProvider.PrimaryFont.GposTable;  // Force load now - thread-safe via TableLoader

            _fontProvider = fontProvider;
            _primaryFont = fontProvider.PrimaryFont;

            // Initialize processors with primary font
            _kerningProvider = new KerningProvider(_primaryFont);
            _ligatureProcessor = new LigatureProcessor(_primaryFont);
            _markToBaseProvider = new MarkToBaseProvider(_primaryFont);
            _singleAdjustmentProvider = new SingleAdjustmentProvider(_primaryFont);
            _singleSubstitutionProcessor = new SingleSubstitutionProcessor(_primaryFont);
            _chainingContextualProcessor = new ChainingContextualProcessor(_primaryFont, _singleSubstitutionProcessor, _ligatureProcessor);
        }

        #region Font Tracking API

        /// <summary>
        /// Gets all fonts used in the last shaping operation.
        /// Used for subsetting and PDF embedding when text uses multiple fonts.
        /// </summary>
        public IEnumerable<OpenTypeFont> GetUsedFonts()
        {
            return _usedFonts;
        }

        /// <summary>
        /// Clears font tracking between different texts.
        /// Call this if you're reusing the same TextShaper for multiple unrelated texts.
        /// </summary>
        public void ResetFontTracking()
        {
            _usedFonts.Clear();
            _fontToIdMap.Clear();
        }

        /// <summary>
        /// Gets or registers a font ID for tracking.
        /// Returns: 0 for primary font, 1+ for fallback fonts.
        /// </summary>
        private byte GetOrRegisterFontId(OpenTypeFont font)
        {
            byte fontId;
            if (_fontToIdMap.TryGetValue(font, out fontId))
            {
                return fontId;
            }

            fontId = (byte)_usedFonts.Count;
            _usedFonts.Add(font);
            _fontToIdMap[font] = fontId;
            return fontId;
        }

        #endregion

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

            // Phase 1: Map characters to glyphs (with font fallback support)
            var glyphs = MapToGlyphs(text);

            // Phase 2: Apply GSUB substitutions (if enabled) - ONLY on primary font glyphs
            if (options.ApplySubstitutions && _primaryFont.GsubTable != null)
            {
                glyphs = ApplyGsubSubstitutions(glyphs, options);
            }

            // Phase 3: Apply GPOS positioning (if enabled) - ONLY on primary font glyphs
            if (options.ApplyPositioning)
            {
                ApplyPositioning(glyphs, options);
            }

            // Phase 4: Build result
            var fontUnitsPerEm = BuildFontUnitsPerEm();
            var fontLineHeights = BuildFontLineHeights();
            return new ShapedText
            {
                OriginalText = text,
                Glyphs = glyphs.ToArray(),
                FontUnitsPerEm = fontUnitsPerEm,
                FontLineHeights = fontLineHeights
            };
        }

        /// <summary>
        /// Extracts character widths and returns a new array.
        /// For repeated calls, consider using ExtractCharWidths(text, fontSize, options, targetArray) 
        /// to avoid allocations.
        /// </summary>
        public double[] ExtractCharWidths(string text, float fontSize, ShapingOptions options)
        {
            var charWidths = new double[text.Length];

            if (string.IsNullOrEmpty(text))
            {
                return charWidths;
            }

            ExtractCharWidthsCore(text, fontSize, options, charWidths);
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
                throw new ArgumentException(
                    string.Format("Target array must be at least as large as text length ({0})", text.Length),
                    "targetArray");
            }

            ExtractCharWidthsCore(text, fontSize, options, targetArray);
        }

        /// <summary>
        /// Core implementation that extracts char widths into provided buffer.
        /// OPTIMIZED: Avoids creating ShapedText object and copying glyphs to array.
        /// Works directly with List<ShapedGlyph> for better memory efficiency.
        /// </summary>
        private void ExtractCharWidthsCore(string text, float fontSize, ShapingOptions options, double[] targetArray)
        {
            // Clear only the portion we will use
            Array.Clear(targetArray, 0, text.Length);

            // Phase 1: Map characters to glyphs
            var glyphs = MapToGlyphs(text);

            // Phase 2: Apply GSUB substitutions (if enabled)
            if (options.ApplySubstitutions && _primaryFont.GsubTable != null)
            {
                glyphs = ApplyGsubSubstitutions(glyphs, options);
            }

            // Phase 3: Apply GPOS positioning (if enabled)
            if (options.ApplyPositioning)
            {
                ApplyPositioning(glyphs, options);
            }

            // Phase 4: Extract widths - must handle multi-font glyphs
            foreach (var glyph in glyphs)
            {
                int charIndex = glyph.ClusterIndex;
                if (charIndex >= 0 && charIndex < text.Length)
                {
                    // Get the font for this glyph
                    OpenTypeFont font = _usedFonts[glyph.FontId];
                    double scaleFactor = fontSize / font.HeadTable.UnitsPerEm;

                    targetArray[charIndex] += glyph.XAdvance * scaleFactor;
                }
            }
        }

        #endregion

        #region Phase 1: Character to Glyph Mapping

        /// <summary>
        /// Maps characters to glyphs using the cmap table.
        /// CORRECTLY handles surrogate pairs for emoji and supplementary plane characters.
        /// Supports multi-font fallback via IFontProvider.
        /// </summary>
        private List<ShapedGlyph> MapToGlyphs(string text)
        {
            var glyphs = new List<ShapedGlyph>(text.Length);

            int i = 0;
            while (i < text.Length)
            {
                uint codePoint;
                int charCount;

                // Check if this is a surrogate pair
                if (i < text.Length - 1 && char.IsHighSurrogate(text[i]))
                {
                    // Potential surrogate pair: 2 chars → 1 Unicode code point
                    char high = text[i];
                    char low = text[i + 1];

                    if (char.IsLowSurrogate(low))
                    {
                        // Valid pair - convert to code point
                        codePoint = (uint)char.ConvertToUtf32(high, low);
                        charCount = 2;
                    }
                    else
                    {
                        // Invalid surrogate pair - treat as .notdef and skip high surrogate
                        codePoint = 0;
                        charCount = 1;
                    }
                }
                else if (char.IsSurrogate(text[i]))
                {
                    // Lone surrogate (invalid) - treat as .notdef
                    codePoint = 0;
                    charCount = 1;
                }
                else
                {
                    // Normal BMP character
                    codePoint = text[i];
                    charCount = 1;
                }

                // Use font provider to find glyph (with fallback support)
                OpenTypeFont font;
                ushort glyphId;
                _fontProvider.TryGetGlyphFont(codePoint, out font, out glyphId);

                // Get font ID for multi-font tracking
                byte fontId = GetOrRegisterFontId(font);

                // Get advance width from the font that contains this glyph
                var hmtxTable = font.HmtxTable;
                var baseAdvance = (short)hmtxTable.GetAdvanceWidth(glyphId);

                glyphs.Add(new ShapedGlyph
                {
                    GlyphId = glyphId,
                    BaseAdvance = baseAdvance,
                    XAdvance = baseAdvance,
                    YAdvance = 0,
                    XOffset = 0,
                    YOffset = 0,
                    ClusterIndex = (ushort)i,
                    CharCount = (byte)charCount,
                    FontId = fontId  // Track which font this glyph comes from
                });

                i += charCount;
            }

            return glyphs;
        }

        #endregion

        #region Phase 2: GSUB Substitutions

        /// <summary>
        /// Applies GSUB substitutions (ligatures, contextual alternates, etc.).
        /// IMPORTANT: Only processes glyphs from primary font (FontId == 0).
        /// Fallback font glyphs are not affected by substitutions.
        /// </summary>
        private List<ShapedGlyph> ApplyGsubSubstitutions(List<ShapedGlyph> glyphs, ShapingOptions options)
        {
            // Quick check: If all glyphs are from primary font, process directly
            bool hasNonPrimaryGlyphs = false;
            foreach (var g in glyphs)
            {
                if (g.FontId != 0)
                {
                    hasNonPrimaryGlyphs = true;
                    break;
                }
            }

            if (!hasNonPrimaryGlyphs)
            {
                // All glyphs are primary - process directly (optimization)
                glyphs = ApplyGsubSubstitutionsInternal(glyphs, options);
                return glyphs;
            }

            // Mixed fonts: Extract primary glyphs, process, then merge back
            var primaryGlyphs = new List<ShapedGlyph>();
            var primaryIndices = new List<int>();

            for (int i = 0; i < glyphs.Count; i++)
            {
                if (glyphs[i].FontId == 0)
                {
                    primaryGlyphs.Add(glyphs[i]);
                    primaryIndices.Add(i);
                }
            }

            if (primaryGlyphs.Count == 0)
            {
                return glyphs; // No primary font glyphs to process
            }

            // Process primary glyphs
            primaryGlyphs = ApplyGsubSubstitutionsInternal(primaryGlyphs, options);

            // Merge back: Replace primary glyphs in original list
            var result = new List<ShapedGlyph>(glyphs);

            // Remove old primary glyphs (in reverse to maintain indices)
            for (int i = primaryIndices.Count - 1; i >= 0; i--)
            {
                result.RemoveAt(primaryIndices[i]);
            }

            // Insert processed primary glyphs at first primary position
            int insertPosition = primaryIndices.Count > 0 ? primaryIndices[0] : 0;
            result.InsertRange(insertPosition, primaryGlyphs);

            return result;
        }

        /// <summary>
        /// Internal method that applies GSUB substitutions to a list of glyphs.
        /// Assumes all glyphs are from the same font.
        /// </summary>
        private List<ShapedGlyph> ApplyGsubSubstitutionsInternal(List<ShapedGlyph> glyphs, ShapingOptions options)
        {
            // Phase 1: Single Substitution (Type 1)
            if (options.GsubFeatures != null && options.GsubFeatures.Count > 0)
            {
                glyphs = _singleSubstitutionProcessor.ApplySubstitutions(glyphs, options.GsubFeatures);
            }

            // Phase 2: Chaining Contextual Substitution (Type 6)
            if (options.GsubFeatures != null && options.GsubFeatures.Contains("liga"))
            {
                glyphs = _chainingContextualProcessor.ApplyContextualSubstitutions(glyphs, "liga");
            }

            // Phase 3: Simple Ligatures (Type 4)
            if (options.GsubFeatures != null && options.GsubFeatures.Contains("liga"))
            {
                _ligatureProcessor.ApplyLigaturesInPlace(glyphs);
            }

            return glyphs;
        }

        #endregion

        #region Phase 3: Positioning

        /// <summary>
        /// Applies positioning adjustments (kerning, mark positioning, etc.).
        /// IMPORTANT: Only processes glyphs from primary font (FontId == 0).
        /// Order matters: Single adjustments → Kerning → Mark positioning
        /// </summary>
        private void ApplyPositioning(List<ShapedGlyph> glyphs, ShapingOptions options)
        {
            if (!options.ApplyPositioning)
            {
                return;
            }

            bool applyAllFeatures = options.GposFeatures == null || options.GposFeatures.Count == 0;

            // Phase 1: Single Adjustment (GPOS Type 1) - primary font only
            ApplySingleAdjustment(glyphs, options);

            // Phase 2: Kerning (GPOS Type 2 / kern table) - primary font only
            if (applyAllFeatures || (options.GposFeatures != null && options.GposFeatures.Contains("kern")))
            {
                ApplyKerning(glyphs);
            }

            // Phase 3: Mark-to-Base positioning (GPOS Type 4) - primary font only
            _markToBaseProvider.ApplyMarkPositioning(glyphs);
        }

        /// <summary>
        /// Applies single glyph adjustments from GPOS Lookup Type 1.
        /// Only applies to primary font glyphs (FontId == 0).
        /// </summary>
        private void ApplySingleAdjustment(List<ShapedGlyph> glyphs, ShapingOptions options)
        {
            List<string> features = options.GposFeatures ?? new List<string>();

            for (int i = 0; i < glyphs.Count; i++)
            {
                // Skip fallback font glyphs
                if (glyphs[i].FontId != 0)
                    continue;

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
        /// Only kerns between primary font glyphs (FontId == 0).
        /// </summary>
        private void ApplyKerning(List<ShapedGlyph> glyphs)
        {
            for (int i = 1; i < glyphs.Count; i++)
            {
                // Only kern if BOTH glyphs are from primary font
                if (glyphs[i - 1].FontId != 0 || glyphs[i].FontId != 0)
                    continue;

                ushort leftGlyph = glyphs[i - 1].GlyphId;
                ushort rightGlyph = glyphs[i].GlyphId;

                short kernValue = _kerningProvider.GetKerning(leftGlyph, rightGlyph);

                if (kernValue != 0)
                {
                    var glyph = glyphs[i - 1];
                    glyph.XAdvance += kernValue;
                    glyphs[i - 1] = glyph;
                }
            }
        }

        #endregion

        #region Utilities

        /// <summary>
        /// Measures the width of text in PDF points.
        /// </summary>
        public float MeasureTextInPoints(string text, float fontSize, ShapingOptions options = null)
        {
            var shaped = Shape(text, options);
            return shaped.GetWidthInPoints(fontSize);
        }

        /// <summary>
        /// Measures the width of text in pixels.
        /// </summary>
        public float MeasureTextInPixels(string text, float fontSize, float dpi, ShapingOptions options = null)
        {
            var shaped = Shape(text, options);
            return shaped.GetWidthInPixels(fontSize, dpi);
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

            float maxWidth = 0;
            foreach (var line in shapedLines)
            {
                float lineWidth = line.GetWidthInPoints(fontSize);
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

        #endregion

        #region Light Shaping Pipeline

        /// <summary>
        /// Shapes text into lightweight GlyphWidth structs optimized for text measurement.
        /// </summary>
        public GlyphWidth[] ShapeLight(string text, ShapingOptions options = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new GlyphWidth[0];
            }

            if (options == null)
            {
                options = ShapingOptions.Default;
            }

            var glyphs = MapToGlyphs(text);

            if (options.ApplySubstitutions && _primaryFont.GsubTable != null)
            {
                glyphs = ApplyGsubSubstitutions(glyphs, options);
            }

            if (options.ApplyPositioning)
            {
                ApplyKerningOnly(glyphs);
            }

            return ExtractGlyphWidths(glyphs);
        }

        private void ApplyKerningOnly(List<ShapedGlyph> glyphs)
        {
            for (int i = 1; i < glyphs.Count; i++)
            {
                // Only kern primary font glyphs
                if (glyphs[i - 1].FontId != 0 || glyphs[i].FontId != 0)
                    continue;

                ushort leftGlyph = glyphs[i - 1].GlyphId;
                ushort rightGlyph = glyphs[i].GlyphId;

                short kernValue = _kerningProvider.GetKerning(leftGlyph, rightGlyph);

                if (kernValue != 0)
                {
                    var glyph = glyphs[i - 1];
                    glyph.XAdvance += kernValue;
                    glyphs[i - 1] = glyph;
                }
            }
        }

        private GlyphWidth[] ExtractGlyphWidths(List<ShapedGlyph> glyphs)
        {
            var result = new GlyphWidth[glyphs.Count];

            for (int i = 0; i < glyphs.Count; i++)
            {
                var g = glyphs[i];
                result[i] = new GlyphWidth
                {
                    XAdvance = (ushort)g.XAdvance,
                    ClusterIndex = g.ClusterIndex,
                    CharCount = g.CharCount
                };
            }

            return result;
        }

        #endregion

        #region Font Metrics

        /// <summary>
        /// Gets single line spacing (baseline-to-baseline distance).
        /// </summary>
        public float GetLineHeightInPoints(float fontSize)
        {
            if (_primaryFont.Os2Table.UseTypoMetrics)
            {
                var typoAscent = _primaryFont.Os2Table.sTypoAscender;
                var typoDescent = _primaryFont.Os2Table.sTypoDescender;
                var typoLineGap = _primaryFont.Os2Table.sTypoLineGap;
                float em = _primaryFont.HeadTable.UnitsPerEm;
                float lineHeight = typoAscent - typoDescent + typoLineGap;
                return (lineHeight / em) * fontSize;
            }
            else
            {
                return GetFontHeightInPoints(fontSize);
            }
        }

        /// <summary>
        /// Calculates the total height of the font, in points.
        /// </summary>
        public float GetFontHeightInPoints(float fontSize)
        {
            var ascent = _primaryFont.Os2Table.usWinAscent;
            var descent = _primaryFont.Os2Table.usWinDescent;
            var em = _primaryFont.HeadTable.UnitsPerEm;

            return (ascent + descent) * (fontSize / em);
        }

        /// <summary>
        /// Calculates the distance from the top of the font's bounding box to the baseline.
        /// </summary>
        /// <param name="fontSize">The font size, in points, for which to calculate the baseline position. Must be a positive value.</param>
        /// <returns>The distance, in points, from the top of the font's bounding box to the baseline for the given font size.</returns>
        public float GetAscentInPoints(float fontSize)
        {
            var ascent = _primaryFont.Os2Table.UseTypoMetrics
                ? (float)_primaryFont.Os2Table.sTypoAscender
                : _primaryFont.Os2Table.usWinAscent;

            var em = _primaryFont.HeadTable.UnitsPerEm;
            return ascent * (fontSize / em);
        }

        /// <summary>
        /// Calculates the font descent in points.
        /// </summary>
        public float GetDescentInPoints(float fontSize)
        {
            var descent = _primaryFont.Os2Table.UseTypoMetrics
                ? (float)Math.Abs(_primaryFont.Os2Table.sTypoDescender)
                : _primaryFont.Os2Table.usWinDescent;

            var em = _primaryFont.HeadTable.UnitsPerEm;
            return descent * (fontSize / em);
        }

        #endregion

        /// <summary>
        /// Builds a UnitsPerEm lookup array indexed by FontId.
        /// Must be called after shaping when _usedFonts is populated.
        /// </summary>
        private ushort[] BuildFontUnitsPerEm()
        {
            if (_usedFonts.Count == 0)
                return new ushort[] { _primaryFont.HeadTable.UnitsPerEm };

            var result = new ushort[_usedFonts.Count];
            for (int i = 0; i < _usedFonts.Count; i++)
            {
                result[i] = _usedFonts[i].HeadTable.UnitsPerEm;
            }
            return result;
        }

        /// <summary>
        /// Builds a line height lookup array (in design units) indexed by FontId.
        /// Uses the same metric selection logic as GetLineHeightInPoints:
        /// if USE_TYPO_METRICS is set, uses sTypoAscender - sTypoDescender + sTypoLineGap;
        /// otherwise uses usWinAscent + usWinDescent.
        /// </summary>
        private int[] BuildFontLineHeights()
        {
            if (_usedFonts.Count == 0)
                return new int[] { GetLineHeightDesignUnits(_primaryFont) };

            var result = new int[_usedFonts.Count];
            for (int i = 0; i < _usedFonts.Count; i++)
            {
                result[i] = GetLineHeightDesignUnits(_usedFonts[i]);
            }
            return result;
        }

        /// <summary>
        /// Gets the line height in design units for a font, using the same
        /// metric selection as GetLineHeightInPoints.
        /// </summary>
        private static int GetLineHeightDesignUnits(OpenTypeFont font)
        {
            if (font.Os2Table.UseTypoMetrics)
            {
                return font.Os2Table.sTypoAscender
                     - font.Os2Table.sTypoDescender
                     + font.Os2Table.sTypoLineGap;
            }
            else
            {
                return font.Os2Table.usWinAscent
                     + font.Os2Table.usWinDescent;
            }
        }
    }
}