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
  03/16/2026         EPPlus Software AB           Multi-font aware width calculation via FontUnitsPerEm
 *************************************************************************************************/
using System;
using System.Diagnostics;
using System.Text;

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Result of text shaping operation containing positioned glyphs.
    /// Supports multi-font text (e.g., primary font + emoji fallback) where glyphs
    /// may originate from fonts with different UnitsPerEm values.
    /// </summary>
    [DebuggerDisplay("Glyphs length: {Glyphs.Length}, OriginalText: {OriginalText}")]
    public class ShapedText
    {
        /// <summary>
        /// Array of shaped glyphs with positioning information.
        /// </summary>
        public ShapedGlyph[] Glyphs { get; set; }

        /// <summary>
        /// The original input text that was shaped.
        /// </summary>
        public string OriginalText { get; set; }

        /// <summary>
        /// UnitsPerEm indexed by FontId. Set by TextShaper after shaping.
        /// FontUnitsPerEm[0] = first used font's UPM, [1] = second font's UPM, etc.
        /// Note: FontId 0 is the first font that was actually used during shaping,
        /// which may be a fallback font if the text starts with e.g. emoji characters.
        /// </summary>
        public ushort[] FontUnitsPerEm { get; set; }

        /// <summary>
        /// Line height per FontId in design units (ascent - descent + lineGap, or winAscent + winDescent).
        /// Set by TextShaper after shaping using the same metric selection logic as
        /// <see cref="TextShaper.GetLineHeightInPoints"/>.
        /// Used by <see cref="GetLineHeightInPoints"/> to compute the correct line height
        /// when glyphs come from fonts with different vertical metrics.
        /// </summary>
        public int[] FontLineHeights { get; set; }

        /// <summary>
        /// Total horizontal advance width in font design units.
        /// This is the sum of all glyph XAdvance values.
        /// Suitable for single-font comparisons (e.g. verifying kerning).
        /// For multi-font text where glyphs may have different UnitsPerEm,
        /// use <see cref="GetWidthInPoints(float)"/> instead.
        /// </summary>
        public int TotalAdvanceWidth
        {
            get
            {
                int total = 0;
                if (Glyphs != null)
                {
                    foreach (var glyph in Glyphs)
                    {
                        total += glyph.XAdvance;
                    }
                }
                return total;
            }
        }

        /// <summary>
        /// Convert advance width to PDF points.
        /// Handles multi-font text correctly by using each glyph's FontId to look up
        /// the correct UnitsPerEm from <see cref="FontUnitsPerEm"/>.
        /// </summary>
        /// <param name="fontSize">Font size in points</param>
        /// <returns>Width in PDF points</returns>
        public float GetWidthInPoints(float fontSize)
        {
            if (Glyphs == null || Glyphs.Length == 0)
                return 0f;

            if (FontUnitsPerEm == null || FontUnitsPerEm.Length == 0)
                return 0f;

            // Fast path: single font — all glyphs share the same UnitsPerEm
            if (FontUnitsPerEm.Length == 1)
            {
                int total = 0;
                foreach (var glyph in Glyphs)
                    total += glyph.XAdvance;
                return (total / (float)FontUnitsPerEm[0]) * fontSize;
            }

            // Multi-font path: convert each glyph individually
            float totalWidth = 0f;
            foreach (var glyph in Glyphs)
            {
                float upm = glyph.FontId < FontUnitsPerEm.Length
                    ? FontUnitsPerEm[glyph.FontId]
                    : FontUnitsPerEm[0];
                totalWidth += (glyph.XAdvance / upm) * fontSize;
            }
            return totalWidth;
        }

        /// <summary>
        /// Gets the line height (baseline-to-baseline distance) in points.
        /// For multi-font text, returns the maximum line height across all fonts
        /// used in this shaped text, ensuring the line is tall enough for every glyph.
        /// </summary>
        /// <param name="fontSize">Font size in points</param>
        /// <returns>Line height in points</returns>
        public float GetLineHeightInPoints(float fontSize)
        {
            if (FontLineHeights == null || FontLineHeights.Length == 0 ||
                FontUnitsPerEm == null || FontUnitsPerEm.Length == 0)
                return fontSize;

            // Single font — fast path
            if (FontLineHeights.Length == 1)
                return (FontLineHeights[0] / (float)FontUnitsPerEm[0]) * fontSize;

            // Multi-font — find which fonts are actually used, return max line height
            float maxLineHeight = 0f;
            var usedFontIds = new System.Collections.Generic.HashSet<byte>();
            foreach (var glyph in Glyphs)
            {
                if (usedFontIds.Add(glyph.FontId) &&
                    glyph.FontId < FontLineHeights.Length &&
                    glyph.FontId < FontUnitsPerEm.Length)
                {
                    float lh = (FontLineHeights[glyph.FontId] / (float)FontUnitsPerEm[glyph.FontId]) * fontSize;
                    if (lh > maxLineHeight)
                        maxLineHeight = lh;
                }
            }
            return maxLineHeight > 0f ? maxLineHeight : fontSize;
        }

        /// <summary>
        /// Convert advance width to pixels. Multi-font aware.
        /// </summary>
        /// <param name="fontSize">Font size in points</param>
        /// <param name="dpi">Screen DPI (typically 96 or 72)</param>
        /// <returns>Width in pixels</returns>
        public float GetWidthInPixels(float fontSize, float dpi)
        {
            return GetWidthInPoints(fontSize) * (dpi / 72f);
        }

        public ShapedText()
        {
            Glyphs = new ShapedGlyph[0];
            OriginalText = string.Empty;
        }
    }
}