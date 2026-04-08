
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
  03/16/2026         EPPlus Software AB           Multi-font aware width calculation
  03/18/2026         EPPlus Software AB           Refactored to use ShapedTextBase
 *************************************************************************************************/
using System.Diagnostics;

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Result of text shaping operation containing positioned glyphs.
    /// Supports multi-font text (e.g., primary font + emoji fallback) where glyphs
    /// may originate from fonts with different UnitsPerEm values.
    /// </summary>
    [DebuggerDisplay("Glyphs length: {Glyphs.Length}, OriginalText: {OriginalText}")]
    public class ShapedText : ShapedTextBase
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
        /// Total horizontal advance width in font design units.
        /// Suitable for single-font comparisons (e.g. verifying kerning).
        /// For multi-font text, use <see cref="ShapedTextBase.GetWidthInPoints"/> instead.
        /// </summary>
        public int TotalAdvanceWidth
        {
            get
            {
                int total = 0;
                if (Glyphs != null)
                {
                    foreach (var glyph in Glyphs)
                        total += glyph.XAdvance;
                }
                return total;
            }
        }

        protected override int GetGlyphCount() =>
            Glyphs?.Length ?? 0;

        protected override int GetGlyphXAdvance(int index) =>
            Glyphs[index].XAdvance;

        protected override byte GetGlyphFontId(int index) =>
            Glyphs[index].FontId;

        public ShapedText()
        {
            Glyphs = new ShapedGlyph[0];
            OriginalText = string.Empty;
        }
    }
}