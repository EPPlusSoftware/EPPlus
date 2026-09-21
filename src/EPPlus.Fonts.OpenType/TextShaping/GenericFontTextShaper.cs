/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/31/2026         EPPlus Software AB           Initial implementation
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.GenericFontWidths;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping
{
    /// <summary>
    /// An <see cref="ITextShaper"/> implementation backed by the serialized font metrics
    /// (.fmtr files in Resources/TextMetrics.zip) instead of a real font file.
    ///
    /// This shaper requires no access to fonts installed on the system, which makes it
    /// usable as a fallback for text measurement and line wrapping when the actual font
    /// cannot be loaded. It is NOT a substitute for the OpenType based
    /// <see cref="TextShaper"/> when real shaping is required (PDF export), because the
    /// underlying data contains no glyph outlines, no glyph ids and no OpenType layout
    /// tables.
    ///
    /// Deliberate limitations, all of them consequences of the .fmtr format:
    /// <list type="bullet">
    /// <item>No GSUB/GPOS. Ligatures, kerning, contextual alternates and mark positioning
    /// are not applied. <see cref="ShapingOptions"/> is accepted but ignored.</item>
    /// <item>No glyph ids. <see cref="ShapedGlyph.GlyphId"/> carries the character code so
    /// that shaped output remains diagnosable; it must not be used for subsetting or
    /// embedding.</item>
    /// <item>No font fallback. All glyphs report FontId 0.</item>
    /// <item>Character advances are quantized into width classes (16 or 32 depending on the
    /// font), so per character advances carry an error of up to half a class width. See
    /// the remarks on <see cref="ToDesignUnits"/>.</item>
    /// <item>Only the Basic Multilingual Plane is covered. The metrics are keyed on
    /// <see cref="char"/>, so characters above U+FFFF get the default width class.</item>
    /// </list>
    ///
    /// The scale factors in <see cref="FontScaleFactors"/> are intentionally NOT applied.
    /// Those exist to align AutoFitColumns with the Excel GUI and are not font metrics.
    /// The same reasoning excludes the digit and East Asian scaling factors used by
    /// GenericFontMetricsTextMeasurerBase. This shaper reports the metrics as they are.
    /// </summary>
    internal class GenericFontTextShaper : ITextShaper
    {
        /// <summary>
        /// Units per em for the virtual font this shaper represents. The .fmtr data has no
        /// units per em of its own, so a value is chosen here and all advances are expressed
        /// relative to it. 1000 matches the fallback used by <see cref="TextShaper"/>.
        /// </summary>
        private const ushort GENERIC_UNITS_PER_EM = 1000;

        /// <summary>
        /// Widths in the .fmtr files are stored as pixels per point of font size, i.e. the
        /// em value already multiplied by 96/72 by the font-labs exporter. This constant
        /// reverses that so widths can be converted to design units.
        /// </summary>
        private const float PixelsPerPointToEm = 72f / 96f;

        private readonly SerializedFontMetrics _metrics;
        private readonly uint _fontKey;
        private readonly ushort _defaultAdvance;
        private readonly ushort _lineHeightDesignUnits;

        /// <summary>
        /// Creates a shaper for the supplied metrics.
        /// </summary>
        internal GenericFontTextShaper(SerializedFontMetrics metrics)
        {
            if (metrics == null)
            {
                throw new ArgumentNullException("metrics");
            }

            _metrics = metrics;
            _fontKey = metrics.GetKey();

            _defaultAdvance = ToDesignUnits(_metrics.DefaultWidth);
            _lineHeightDesignUnits = ToDesignUnits(_metrics.LineHeight1em);
        }

        /// <summary>
        /// Attempts to create a shaper for a font family and style. Returns false when the
        /// family has no serialized metrics, in which case the caller should fall back.
        /// </summary>
        internal static bool TryCreate(string fontFamily, MeasurementFontStyles style, out GenericFontTextShaper shaper)
        {
            shaper = null;
            if (string.IsNullOrEmpty(fontFamily))
            {
                return false;
            }

            // ResolveKey rather than GetKey so a missing subfamily falls back to the family's
            // Regular, matching what the measurer does.
            var fontKey = GenericTextMeasurerKey.ResolveKey(fontFamily, style);
            if (fontKey == uint.MaxValue)
            {
                return false;
            }

            var metrics = GenericFontMetricsCache.GetMetrics(fontKey);
            if (metrics == null)
            {
                return false;
            }

            shaper = new GenericFontTextShaper(metrics);
            return true;
        }

        /// <summary>
        /// The font key (family and subfamily) these metrics were loaded for.
        /// </summary>
        internal uint FontKey
        {
            get { return _fontKey; }
        }

        public ushort UnitsPerEm
        {
            get { return GENERIC_UNITS_PER_EM; }
        }

        /// <inheritdoc/>
        public bool HasGlyphIds
        {
            get { return false; }
        }

        #region Horizontal shaping

        public ShapedText Shape(string text, ShapingOptions options = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new ShapedText
                {
                    OriginalText = text ?? string.Empty,
                    Glyphs = new ShapedGlyph[0],
                    FontUnitsPerEm = new ushort[] { GENERIC_UNITS_PER_EM },
                    FontLineHeights = new int[] { _lineHeightDesignUnits }
                };
            }

            var glyphs = MapToGlyphs(text);

            return new ShapedText
            {
                OriginalText = text,
                Glyphs = glyphs.ToArray(),
                FontUnitsPerEm = new ushort[] { GENERIC_UNITS_PER_EM },
                FontLineHeights = new int[] { _lineHeightDesignUnits }
            };
        }

        public ShapedLightText ShapeLight(string text, ShapingOptions options = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new ShapedLightText
                {
                    Glyphs = new GlyphWidth[0],
                    FontUnitsPerEm = new ushort[] { GENERIC_UNITS_PER_EM }
                };
            }

            var glyphs = MapToGlyphs(text);
            var result = new GlyphWidth[glyphs.Count];
            for (var i = 0; i < glyphs.Count; i++)
            {
                var g = glyphs[i];
                result[i] = new GlyphWidth
                {
                    XAdvance = (ushort)g.XAdvance,
                    ClusterIndex = g.ClusterIndex,
                    CharCount = g.CharCount,
                    FontId = g.FontId
                };
            }

            return new ShapedLightText
            {
                Glyphs = result,
                FontUnitsPerEm = new ushort[] { GENERIC_UNITS_PER_EM }
            };
        }

        public ShapedText[] ShapeLines(string text, ShapingOptions options = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new ShapedText[0];
            }

            var lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var result = new ShapedText[lines.Length];
            for (var i = 0; i < lines.Length; i++)
            {
                result[i] = Shape(lines[i], options);
            }
            return result;
        }

        /// <summary>
        /// Maps characters to advances. One glyph per character, except that a valid
        /// surrogate pair produces a single glyph with CharCount 2.
        /// </summary>
        private List<ShapedGlyph> MapToGlyphs(string text)
        {
            var glyphs = new List<ShapedGlyph>(text.Length);

            var i = 0;
            while (i < text.Length)
            {
                int charCount;
                ushort advance;
                ushort glyphId;

                if (i < text.Length - 1 && char.IsHighSurrogate(text[i]) && char.IsLowSurrogate(text[i + 1]))
                {
                    // Supplementary plane. The metrics are keyed on char and cover the BMP
                    // only, so there is nothing better available than the default width.
                    charCount = 2;
                    advance = _defaultAdvance;
                    glyphId = 0;
                }
                else if (char.IsSurrogate(text[i]))
                {
                    // Lone surrogate. Zero width, mirroring .notdef handling in TextShaper.
                    charCount = 1;
                    advance = 0;
                    glyphId = 0;
                }
                else
                {
                    charCount = 1;
                    advance = GetAdvance(text[i]);
                    glyphId = text[i];
                }

                glyphs.Add(new ShapedGlyph
                {
                    GlyphId = glyphId,
                    BaseAdvance = (short)advance,
                    XAdvance = (short)advance,
                    YAdvance = 0,
                    XOffset = 0,
                    YOffset = 0,
                    ClusterIndex = (ushort)i,
                    CharCount = (byte)charCount,
                    FontId = 0
                });

                i += charCount;
            }

            return glyphs;
        }

        /// <summary>
        /// Resolves the advance width of a single BMP character, in design units.
        /// </summary>
        private ushort GetAdvance(char c)
        {
            // Control characters carry no width. GenericFontMetricsTextMeasurerBase makes the
            // same exclusion; note that this also covers CR and LF, so a caller that shapes
            // multi-line text through Shape() rather than ShapeLines() gets zero width line
            // breaks instead of default width ones.
            if (char.IsControl(c))
            {
                return 0;
            }

            if (IsEastAsianChar(c))
            {
                return GetEastAsianAdvance(c);
            }

            return ToDesignUnits(_metrics.GetCharacterWidth(c));
        }

        /// <summary>
        /// East Asian characters are full width (one em) regardless of the font, with the
        /// half width Katakana block at half an em.
        ///
        /// Unlike GenericFontMetricsTextMeasurerBase this applies neither the 1.13 Kanji
        /// scaling factor nor the 1.05 bold factor. Both are Excel GUI calibration rather
        /// than font metrics and belong in the measurer, not here.
        /// </summary>
        private static ushort GetEastAsianAdvance(char c)
        {
            var cc = (int)c;
            // U+FF61 - U+FF9F, half width Katakana and punctuation.
            if (cc >= 0xFF61 && cc <= 0xFF9F)
            {
                return (ushort)(GENERIC_UNITS_PER_EM / 2);
            }
            return GENERIC_UNITS_PER_EM;
        }

        /// <summary>
        /// Returns true when the character falls inside one of the Japanese/Kanji ranges.
        /// Uses an explicit loop rather than LINQ; this runs once per character and the
        /// LINQ version in GenericFontMetricsTextMeasurerBase allocates an enumerator and
        /// a closure on every call.
        /// </summary>
        private static bool IsEastAsianChar(char c)
        {
            var cc = (int)c;

            // Cheap rejection of Latin and most of the BMP below the CJK blocks.
            if (cc < 0x2E80)
            {
                return false;
            }

            foreach (var range in UniCodeRange.JapaneseKanji)
            {
                if (range.IsInRange(cc))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Converts a width from the .fmtr format (pixels per point of font size) to design
        /// units of the virtual font.
        ///
        /// Rounding to whole design units adds an error below 0.001 em, which is negligible
        /// next to the quantization already present in the source data: a class width step
        /// is 0.114 em for the 16 class fonts (Calibri, Arial) and 0.063 - 0.075 em for the
        /// 32 class fonts (Aptos Narrow, Segoe UI, Tahoma). At 11pt that is 1.67px and
        /// 0.92 - 1.10px respectively, so a single character advance can be off by half of
        /// that. The errors partly cancel across a string when measuring, but a caller that
        /// positions character by character accumulates them.
        /// </summary>
        private static ushort ToDesignUnits(float fmtrWidth)
        {
            if (fmtrWidth <= 0f)
            {
                return 0;
            }
            var designUnits = Math.Round(fmtrWidth * PixelsPerPointToEm * GENERIC_UNITS_PER_EM,
                                         MidpointRounding.AwayFromZero);
            if (designUnits > ushort.MaxValue)
            {
                return ushort.MaxValue;
            }
            return (ushort)designUnits;
        }

        #endregion

        #region Vertical shaping

        /// <summary>
        /// Shapes text for vertical layout. The .fmtr format has no vertical metrics, so the
        /// horizontal advance is used as the advance height. This mirrors what
        /// <see cref="TextShaper"/> does for fonts without a vmtx table, but it gives poor
        /// stacking for narrow Latin characters and should be revisited if vertical text is
        /// actually routed through this shaper.
        /// </summary>
        public ShapedVerticalText ShapeVertical(string text, ShapingOptions options = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new ShapedVerticalText
                {
                    OriginalText = text ?? string.Empty,
                    Glyphs = new VerticalShapedGlyph[0]
                };
            }

            var horizontal = MapToGlyphs(text);
            var glyphs = new VerticalShapedGlyph[horizontal.Count];
            for (var i = 0; i < horizontal.Count; i++)
            {
                var g = horizontal[i];
                var advance = (ushort)g.XAdvance;
                glyphs[i] = new VerticalShapedGlyph(
                    g.GlyphId,
                    advance,   // advanceHeight, no vertical metrics available
                    0,         // topSideBearing
                    advance,   // advanceWidth, used for centering
                    g.ClusterIndex,
                    g.CharCount,
                    0);
            }

            return new ShapedVerticalText
            {
                OriginalText = text,
                Glyphs = glyphs
            };
        }

        public VerticalGlyphHeight[] ShapeLightVertical(string text, ShapingOptions options = null)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new VerticalGlyphHeight[0];
            }

            var horizontal = MapToGlyphs(text);
            var result = new VerticalGlyphHeight[horizontal.Count];
            for (var i = 0; i < horizontal.Count; i++)
            {
                var g = horizontal[i];
                result[i] = new VerticalGlyphHeight
                {
                    YAdvance = (ushort)g.XAdvance,
                    ClusterIndex = g.ClusterIndex,
                    CharCount = g.CharCount
                };
            }
            return result;
        }

        #endregion

        #region Character widths

        public double[] ExtractCharWidths(string text, float fontSize, ShapingOptions options)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new double[text == null ? 0 : text.Length];
            }

            var charWidths = new double[text.Length];
            ExtractCharWidthsCore(text, fontSize, charWidths);
            return charWidths;
        }

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

            ExtractCharWidthsCore(text, fontSize, targetArray);
        }

        private void ExtractCharWidthsCore(string text, float fontSize, double[] targetArray)
        {
            Array.Clear(targetArray, 0, text.Length);

            var scaleFactor = fontSize / (double)GENERIC_UNITS_PER_EM;
            var glyphs = MapToGlyphs(text);

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

        #region Font metrics

        public float GetLineHeightInPoints(float fontSize)
        {
            return _metrics.LineHeight1em * PixelsPerPointToEm * fontSize;
        }

        /// <summary>
        /// Total font height, ascent plus descent, excluding the line gap. This is the same
        /// distinction TextShaper makes: GetLineHeightInPoints includes the leading between
        /// lines, this does not.
        /// </summary>
        public float GetFontHeightInPoints(float fontSize)
        {
            return (_metrics.Ascender1em + _metrics.Descender1em) * PixelsPerPointToEm * fontSize;
        }

        /// <summary>
        /// Distance from the top of the line box down to the baseline.
        ///
        /// Exact for version 2 metrics. For version 1 files the value is split out of the line
        /// height by a fixed ratio in SerializedFontMetrics, which is out by up to 7% of the
        /// font height for the extremes of the shipped library.
        /// </summary>
        public float GetAscentInPoints(float fontSize)
        {
            return _metrics.Ascender1em * PixelsPerPointToEm * fontSize;
        }

        /// <summary>
        /// Distance from the baseline down to the bottom of the line box. See
        /// <see cref="GetAscentInPoints"/> for the version 1 caveat.
        /// </summary>
        public float GetDescentInPoints(float fontSize)
        {
            return _metrics.Descender1em * PixelsPerPointToEm * fontSize;
        }

        #endregion
    }
}