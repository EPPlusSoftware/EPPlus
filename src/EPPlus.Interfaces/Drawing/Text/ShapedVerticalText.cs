/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/19/2026         EPPlus Software AB           Vertical text shaping support
 *************************************************************************************************/
using System.Diagnostics;

namespace OfficeOpenXml.Interfaces.Drawing.Text
{
    /// <summary>
    /// Result of a vertical text shaping operation containing positioned glyphs.
    /// Used for Excel vertical text mode (text rotation value 255).
    /// Analogous to <see cref="ShapedText"/> for horizontal text.
    /// </summary>
    [DebuggerDisplay("Glyphs length: {Glyphs.Length}, OriginalText: {OriginalText}")]
    public class ShapedVerticalText
    {
        /// <summary>
        /// The original input text that was shaped.
        /// </summary>
        public string OriginalText { get; set; }

        /// <summary>
        /// Array of vertically shaped glyphs with positioning information.
        /// Glyphs are ordered top-to-bottom.
        /// </summary>
        public VerticalShapedGlyph[] Glyphs { get; set; }

        /// <summary>
        /// Total vertical advance height in font design units.
        /// This is the sum of all glyph YAdvance values and represents
        /// the total height of the text column.
        /// </summary>
        public int TotalAdvanceHeight
        {
            get
            {
                int total = 0;
                if (Glyphs != null)
                {
                    foreach (var glyph in Glyphs)
                    {
                        total += glyph.YAdvance;
                    }
                }
                return total;
            }
        }

        /// <summary>
        /// Converts total advance height to points.
        /// </summary>
        /// <param name="fontSize">Font size in points</param>
        /// <param name="unitsPerEm">Units per EM from the font head table</param>
        /// <returns>Height in points</returns>
        public float GetHeightInPoints(float fontSize, float unitsPerEm)
        {
            return (TotalAdvanceHeight / unitsPerEm) * fontSize;
        }

        /// <summary>
        /// Creates a new ShapedVerticalText with default values.
        /// </summary>
        public ShapedVerticalText()
        {
            Glyphs = new VerticalShapedGlyph[0];
            OriginalText = string.Empty;
        }
    }
}