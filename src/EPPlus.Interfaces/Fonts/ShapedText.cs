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
 *************************************************************************************************/
using System;
using System.Diagnostics;
using System.Text;

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Result of text shaping operation containing positioned glyphs.
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
        /// Total horizontal advance width in font units.
        /// This is the sum of all glyph XAdvance values.
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
        /// Convert total advance width to PDF points.
        /// </summary>
        /// <param name="fontSize">Font size in points</param>
        /// <param name="unitsPerEm">Units per EM from font head table</param>
        /// <returns>Width in PDF points</returns>
        public float GetWidthInPoints(float fontSize, float unitsPerEm)
        {
            return (TotalAdvanceWidth / unitsPerEm) * fontSize;
        }

        /// <summary>
        /// Convert total advance width to pixels.
        /// </summary>
        /// <param name="fontSize">Font size in points</param>
        /// <param name="dpi">Screen DPI (typically 96 or 72)</param>
        /// <param name="unitsPerEm">Units per EM from font head table</param>
        /// <returns>Width in pixels</returns>
        public float GetWidthInPixels(float fontSize, float dpi, float unitsPerEm)
        {
            float points = GetWidthInPoints(fontSize, unitsPerEm);
            return points * (dpi / 72f);
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
                // Show glyph (using glyph ID)
                sb.AppendFormat("<{0:X4}> Tj\n", glyph.GlyphId);

                // Calculate advance in PDF points
                float advanceX = (glyph.XAdvance / unitsPerEm) * fontSize;

                // Apply Y-offset if present (for subscript/superscript)
                if (Math.Abs(glyph.YOffset) > 0.001f)
                {
                    float offsetY = (glyph.YOffset / unitsPerEm) * fontSize;
                    sb.AppendFormat("0 {0:F3} Td\n", offsetY);
                }

                // Advance to next position
                sb.AppendFormat("{0:F3} 0 Td\n", advanceX);
            }

            sb.AppendLine("ET");
            return sb.ToString();
        }

        public ShapedText()
        {
            Glyphs = new ShapedGlyph[0];
            OriginalText = string.Empty;
        }
    }
}