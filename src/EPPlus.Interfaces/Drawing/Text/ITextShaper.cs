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

namespace OfficeOpenXml.Interfaces.Drawing.Text
{
    /// <summary>
    /// Core text shaping - converts text to positioned glyphs.
    /// Works in font design units only.
    /// </summary>
    public interface ITextShaper
    {
        /// <summary>
        /// Shapes text applying OpenType features (ligatures, kerning, etc.).
        /// Returns positioned glyphs in font design units.
        /// </summary>
        ShapedText Shape(string text, ShapingOptions options = null);

        /// <summary>
        /// Shapes text into lightweight GlyphWidth structs optimized for text measurement.
        /// This method is 85% more memory efficient than Shape() and is designed for
        /// text wrapping scenarios where only character widths are needed.
        /// Uses simplified pipeline: character mapping + essential OpenType features only.
        /// </summary>
        /// <param name="text">Text to shape</param>
        /// <param name="options">Shaping options (ligatures and kerning supported)</param>
        /// <returns>Array of lightweight glyph width structs (8 bytes each)</returns>
        GlyphWidth[] ShapeLight(string text, ShapingOptions options = null);

        /// <summary>
        /// Shapes multiple lines (splits on CR/LF/CRLF and shapes each line).
        /// Returns array of shaped lines in font design units.
        /// </summary>
        ShapedText[] ShapeLines(string text, ShapingOptions options = null);

        double[] ExtractCharWidths(string text, float fontSize, ShapingOptions options);

        void ExtractCharWidths(string text, float fontSize, ShapingOptions options, double[] targetArray);

        // === Font Metrics (in design units or converted to points) ===

        /// <summary>
        /// Gets single line spacing (baseline-to-baseline) in points.
        /// </summary>
        float GetLineHeightInPoints(float fontSize);

        /// <summary>
        /// Gets total font height (ascent + descent) in points.
        /// </summary>
        float GetFontHeightInPoints(float fontSize);

        /// <summary>
        /// Gets baseline distance from top of container in points.
        /// </summary>
        float GetBaseLineInPoints(float fontSize);

        /// <summary>
        /// Gets descent distance (below baseline) in points.
        /// </summary>
        float GetDescentInPoints(float fontSize);

        /// <summary>
        /// Gets the font's units per em (for manual conversions).
        /// </summary>
        ushort UnitsPerEm { get; }
    }
}