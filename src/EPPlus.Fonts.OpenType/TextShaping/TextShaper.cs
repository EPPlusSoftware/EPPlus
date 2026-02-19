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
  02/19/2026         EPPlus Software AB           Refactored to partial class, added IFontProvider
                                                  support for fallback fonts and vertical shaping
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
    public partial class TextShaper : ITextShaper
    {
        private readonly OpenTypeFont _primaryFont;
        private readonly IFontProvider _fontProvider;
        private readonly KerningProvider _kerningProvider;
        private readonly LigatureProcessor _ligatureProcessor;
        private readonly MarkToBaseProvider _markToBaseProvider;
        private readonly SingleAdjustmentProvider _singleAdjustmentProvider;
        private readonly SingleSubstitutionProcessor _singleSubstitutionProcessor;
        private readonly ChainingContextualProcessor _chainingContextualProcessor;

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

        public TextShaper(OpenTypeFont font)
            : this(new DefaultFontProvider(font))
        {
        }

        public TextShaper(IFontProvider fontProvider)
        {
            if (fontProvider == null)
                throw new ArgumentNullException("fontProvider");
            if (fontProvider.PrimaryFont == null)
                throw new ArgumentException("Primary font cannot be null in font provider", "fontProvider");

            _fontProvider = fontProvider;
            _primaryFont = fontProvider.PrimaryFont;

            _kerningProvider = new KerningProvider(_primaryFont);
            _ligatureProcessor = new LigatureProcessor(_primaryFont);
            _markToBaseProvider = new MarkToBaseProvider(_primaryFont);
            _singleAdjustmentProvider = new SingleAdjustmentProvider(_primaryFont);
            _singleSubstitutionProcessor = new SingleSubstitutionProcessor(_primaryFont);
            _chainingContextualProcessor = new ChainingContextualProcessor(_primaryFont, _singleSubstitutionProcessor, _ligatureProcessor);

            // Register primary font immediately so FontId 0 is always the primary font
            GetOrRegisterFontId(_primaryFont);
        }

        #region Font Tracking

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
            // Re-register primary font so FontId 0 is always the primary font
            GetOrRegisterFontId(_primaryFont);
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
        public ShapedText Shape(string text)
        {
            return Shape(text, ShapingOptions.Default);
        }

        /// <summary>
        /// Shape text with specified options.
        /// Note: Newline characters (\n, \r, \r\n) are treated as regular characters.
        /// For multi-line text, use ShapeLines() instead.
        /// </summary>
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
            if (options.ApplySubstitutions && _primaryFont.GsubTable != null)
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
        /// </summary>
        private void ExtractCharWidthsCore(string text, float fontSize, ShapingOptions options, double[] targetArray)
        {
            Array.Clear(targetArray, 0, text.Length);

            var glyphs = MapToGlyphs(text);

            if (options.ApplySubstitutions && _primaryFont.GsubTable != null)
            {
                glyphs = ApplyGsubSubstitutions(glyphs, options);
            }

            if (options.ApplyPositioning)
            {
                ApplyPositioning(glyphs, options);
            }

            double scaleFactor = fontSize / UnitsPerEm;

            foreach (var glyph in glyphs)
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
            var cmapTable = _primaryFont.CmapTable;
            var hmtxTable = _primaryFont.HmtxTable;

            for (ushort i = 0; i < text.Length; i++)
            {
                char c = text[i];

                int glyphId = cmapTable.MapCharToGlyph(c);

                if (glyphId < 0)
                {
                    glyphId = 0; // .notdef
                }

                var baseAdvance = (short)hmtxTable.GetAdvanceWidth((ushort)glyphId);

                glyphs.Add(new ShapedGlyph
                {
                    GlyphId = (ushort)glyphId,
                    BaseAdvance = baseAdvance,
                    XAdvance = baseAdvance,
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
            if (options.GsubFeatures != null && options.GsubFeatures.Count > 0)
            {
                glyphs = _singleSubstitutionProcessor.ApplySubstitutions(glyphs, options.GsubFeatures);
            }

            if (options.GsubFeatures != null && options.GsubFeatures.Contains("liga"))
            {
                glyphs = _chainingContextualProcessor.ApplyContextualSubstitutions(glyphs, "liga");
            }

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
        /// </summary>
        private void ApplyPositioning(List<ShapedGlyph> glyphs, ShapingOptions options)
        {
            if (!options.ApplyPositioning)
            {
                return;
            }

            bool applyAllFeatures = options.GposFeatures == null || options.GposFeatures.Count == 0;

            ApplySingleAdjustment(glyphs, options);

            if (applyAllFeatures || (options.GposFeatures != null && options.GposFeatures.Contains("kern")))
            {
                ApplyKerning(glyphs);
            }

            _markToBaseProvider.ApplyMarkPositioning(glyphs);
        }

        private void ApplySingleAdjustment(List<ShapedGlyph> glyphs, ShapingOptions options)
        {
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

        private void ApplyKerning(List<ShapedGlyph> glyphs)
        {
            for (int i = 1; i < glyphs.Count; i++)
            {
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

        /// <summary>
        /// Applies only kerning adjustments for wrapping.
        /// Skips other GPOS features as they don't affect line breaking decisions.
        /// </summary>
        private void ApplyKerningOnly(List<ShapedGlyph> glyphs)
        {
            for (int i = 1; i < glyphs.Count; i++)
            {
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
            float unitsPerEm = _primaryFont.HeadTable.UnitsPerEm;

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

        #endregion

        #region Font Metrics

        /// <summary>
        /// Gets single line spacing (baseline-to-baseline) in points.
        /// </summary>
        public float GetLineHeightInPoints(float fontSize)
        {
            var hhea = _primaryFont.HheaTable;
            float unitsPerEm = _primaryFont.HeadTable.UnitsPerEm;
            int lineHeightUnits = hhea.ascender - hhea.descender + hhea.lineGap;
            return (lineHeightUnits / unitsPerEm) * fontSize;
        }

        /// <summary>
        /// Gets font height (ascent + descent only, no line gap) in points.
        /// </summary>
        public float GetFontHeightInPoints(float fontSize)
        {
            var hhea = _primaryFont.HheaTable;
            float unitsPerEm = _primaryFont.HeadTable.UnitsPerEm;
            int fontHeightUnits = hhea.ascender - hhea.descender;
            return (fontHeightUnits / unitsPerEm) * fontSize;
        }

        /// <summary>
        /// Gets single line spacing (baseline-to-baseline) in points.
        /// Uses typo metrics if USE_TYPO_METRICS flag is set, otherwise uses Win metrics.
        /// </summary>
        public double GetLineHeightInPoints(double fontSize)
        {
            if (_primaryFont.Os2Table.UseTypoMetrics)
            {
                var typoAscent = _primaryFont.Os2Table.sTypoAscender;
                var typoDescent = _primaryFont.Os2Table.sTypoDescender;
                var typoLineGap = _primaryFont.Os2Table.sTypoLineGap;
                double em = _primaryFont.HeadTable.UnitsPerEm;
                double lineHeight = typoAscent - typoDescent + typoLineGap;
                return (lineHeight / em) * fontSize;
            }
            else
            {
                return GetFontHeightInPoints(fontSize);
            }
        }

        /// <summary>
        /// Calculates the total height of the font in points for the specified font size.
        /// </summary>
        public double GetFontHeightInPoints(double fontSize)
        {
            var ascent = _primaryFont.Os2Table.usWinAscent;
            var descent = _primaryFont.Os2Table.usWinDescent;
            var em = _primaryFont.HeadTable.UnitsPerEm;
            return (ascent + descent) * (fontSize / em);
        }

        /// <summary>
        /// Calculates the distance from the top of the font's bounding box to the baseline in points.
        /// </summary>
        public double GetBaseLineInPoints(double fontSize)
        {
            var ascent = _primaryFont.Os2Table.UseTypoMetrics
                ? (double)_primaryFont.Os2Table.sTypoAscender
                : (double)_primaryFont.Os2Table.usWinAscent;

            var em = _primaryFont.HeadTable.UnitsPerEm;
            return ascent * (fontSize / em);
        }

        /// <summary>
        /// Calculates the font descent in points for the specified font size.
        /// </summary>
        public double GetDescentInPoints(double fontSize)
        {
            var descent = _primaryFont.Os2Table.UseTypoMetrics
                ? (double)Math.Abs(_primaryFont.Os2Table.sTypoDescender)
                : _primaryFont.Os2Table.usWinDescent;

            var em = _primaryFont.HeadTable.UnitsPerEm;
            return descent * (fontSize / em);
        }

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
            float unitsPerEm = _primaryFont.HeadTable.UnitsPerEm;
            return shaped.GetWidthInPoints(fontSize, unitsPerEm);
        }

        /// <summary>
        /// Measures the width of text in pixels.
        /// </summary>
        public float MeasureTextInPixels(string text, float fontSize, float dpi, ShapingOptions options = null)
        {
            var shaped = Shape(text, options);
            float unitsPerEm = _primaryFont.HeadTable.UnitsPerEm;
            return shaped.GetWidthInPixels(fontSize, dpi, unitsPerEm);
        }

        #endregion
    }
}