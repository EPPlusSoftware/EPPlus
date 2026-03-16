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
        /// Total horizontal advance width in font design units.
        /// This is the sum of all glyph XAdvance values.
        /// WARNING: When glyphs come from fonts with different UnitsPerEm, this sum
        /// mixes design units from different scales and should NOT be used for point/pixel
        /// conversion. Use <see cref="GetWidthInPoints(float)"/> instead.
        /// This property remains for backward compatibility with single-font usage.
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
        /// This is the preferred overload for all new code.
        /// </summary>
        /// <param name="fontSize">Font size in points</param>
        /// <returns>Width in PDF points</returns>
        public float GetWidthInPoints(float fontSize)
        {
            if (Glyphs == null || Glyphs.Length == 0)
                return 0f;

            // No FontUnitsPerEm set — fall back to simple calculation with default 1000
            if (FontUnitsPerEm == null || FontUnitsPerEm.Length == 0)
                return (TotalAdvanceWidth / 1000f) * fontSize;

            // Fast path: single font — all glyphs share the same UnitsPerEm
            if (FontUnitsPerEm.Length == 1)
                return (TotalAdvanceWidth / (float)FontUnitsPerEm[0]) * fontSize;

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
        /// Convert advance width to pixels. Multi-font aware.
        /// </summary>
        /// <param name="fontSize">Font size in points</param>
        /// <param name="dpi">Screen DPI (typically 96 or 72)</param>
        /// <returns>Width in pixels</returns>
        public float GetWidthInPixels(float fontSize, float dpi)
        {
            return GetWidthInPoints(fontSize) * (dpi / 72f);
        }

        /// <summary>
        /// Generate PDF text operators for rendering this shaped text.
        /// </summary>
        /// <param name="fontSize">Font size in points</param>
        /// <param name="x">Starting X position in PDF coordinates</param>
        /// <param name="y">Starting Y position in PDF coordinates</param>
        /// <param name="unitsPerEm">Units per EM from font head table</param>
        /// <returns>PDF operator string</returns>
        public string ToPdfOperators(float fontSize, float x, float y, float unitsPerEm)
        {
            var sb = new StringBuilder();

            sb.AppendLine("BT");
            sb.AppendFormat("/F1 {0} Tf\n", fontSize);
            sb.AppendFormat("{0} {1} Td\n", x, y);

            foreach (var glyph in Glyphs)
            {
                sb.AppendFormat("<{0:X4}> Tj\n", glyph.GlyphId);

                float glyphUpm = GetUnitsPerEm(glyph.FontId, unitsPerEm);
                float advanceX = (glyph.XAdvance / glyphUpm) * fontSize;

                if (Math.Abs(glyph.YOffset) > 0.001f)
                {
                    float offsetY = (glyph.YOffset / glyphUpm) * fontSize;
                    sb.AppendFormat("0 {0:F3} Td\n", offsetY);
                }

                sb.AppendFormat("{0:F3} 0 Td\n", advanceX);
            }

            sb.AppendLine("ET");
            return sb.ToString();
        }

        /// <summary>
        /// Gets the UnitsPerEm for a given FontId.
        /// Uses FontUnitsPerEm array if available, otherwise falls back to the provided default.
        /// </summary>
        private float GetUnitsPerEm(byte fontId, float fallback)
        {
            if (FontUnitsPerEm != null && fontId < FontUnitsPerEm.Length)
                return FontUnitsPerEm[fontId];
            return fallback;
        }

        public ShapedText()
        {
            Glyphs = new ShapedGlyph[0];
            OriginalText = string.Empty;
        }
    }
}