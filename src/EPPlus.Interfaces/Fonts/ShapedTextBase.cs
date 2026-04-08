/*************************************************************************************************
  Required Notice: Copyright(C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author Change
 *************************************************************************************************
  03/18/2026         EPPlus Software AB           Base class for shaped text results
 *************************************************************************************************/

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Base class for shaped text results. Provides multi-font aware width and
    /// line height calculation shared by both <see cref="ShapedText"/> (full shaping)
    /// and <see cref="ShapedLightText"/> (lightweight measurement).
    /// Subclasses provide glyph data via <see cref="GetGlyphCount"/>,
    /// <see cref="GetGlyphXAdvance"/> and <see cref="GetGlyphFontId"/>.
    /// </summary>
    public abstract class ShapedTextBase
    {
        /// <summary>
        /// UnitsPerEm indexed by FontId. Set by TextShaper after shaping.
        /// </summary>
        public ushort[] FontUnitsPerEm { get; set; }

        /// <summary>
        /// Line height per FontId in design units.
        /// Set by TextShaper after shaping.
        /// </summary>
        public int[] FontLineHeights { get; set; }

        /// <summary>
        /// Number of glyphs in this result.
        /// </summary>
        protected abstract int GetGlyphCount();

        /// <summary>
        /// Gets the XAdvance of the glyph at the specified index.
        /// </summary>
        protected abstract int GetGlyphXAdvance(int index);

        /// <summary>
        /// Gets the FontId of the glyph at the specified index.
        /// </summary>
        protected abstract byte GetGlyphFontId(int index);

        /// <summary>
        /// Convert advance width to PDF points.
        /// Handles multi-font text correctly by using each glyph's FontId to look up
        /// the correct UnitsPerEm.
        /// </summary>
        public float GetWidthInPoints(float fontSize)
        {
            int count = GetGlyphCount();
            if (count == 0 || FontUnitsPerEm == null || FontUnitsPerEm.Length == 0)
                return 0f;

            // Fast path: single font
            if (FontUnitsPerEm.Length == 1)
            {
                int total = 0;
                for (int i = 0; i < count; i++)
                    total += GetGlyphXAdvance(i);
                return (total / (float)FontUnitsPerEm[0]) * fontSize;
            }

            // Multi-font path
            float totalWidth = 0f;
            for (int i = 0; i < count; i++)
            {
                byte fontId = GetGlyphFontId(i);
                float upm = fontId < FontUnitsPerEm.Length
                    ? FontUnitsPerEm[fontId]
                    : FontUnitsPerEm[0];
                totalWidth += (GetGlyphXAdvance(i) / upm) * fontSize;
            }
            return totalWidth;
        }

        /// <summary>
        /// Convert advance width to pixels. Multi-font aware.
        /// </summary>
        public float GetWidthInPixels(float fontSize, float dpi)
        {
            return GetWidthInPoints(fontSize) * (dpi / 72f);
        }

        /// <summary>
        /// Gets the line height (baseline-to-baseline distance) in points.
        /// For multi-font text, returns the maximum line height across all fonts
        /// used, ensuring the line is tall enough for every glyph.
        /// </summary>
        public float GetLineHeightInPoints(float fontSize)
        {
            if (FontLineHeights == null || FontLineHeights.Length == 0 ||
                FontUnitsPerEm == null || FontUnitsPerEm.Length == 0)
                return fontSize;

            if (FontLineHeights.Length == 1)
                return (FontLineHeights[0] / (float)FontUnitsPerEm[0]) * fontSize;

            // Multi-font — max across actually used fonts
            int count = GetGlyphCount();
            float maxLineHeight = 0f;
            uint seenFontIds = 0; // bit field, supports up to 32 fonts (more than enough)

            for (int i = 0; i < count; i++)
            {
                byte fontId = GetGlyphFontId(i);
                uint bit = 1u << fontId;
                if ((seenFontIds & bit) != 0)
                    continue;
                seenFontIds |= bit;

                if (fontId < FontLineHeights.Length && fontId < FontUnitsPerEm.Length)
                {
                    float lh = (FontLineHeights[fontId] / (float)FontUnitsPerEm[fontId]) * fontSize;
                    if (lh > maxLineHeight)
                        maxLineHeight = lh;
                }
            }
            return maxLineHeight > 0f ? maxLineHeight : fontSize;
        }
    }
}