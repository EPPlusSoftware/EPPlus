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

namespace OfficeOpenXml.Interfaces.Fonts
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
        /// Shapes text into a lightweight result optimized for text measurement.
        /// Returns <see cref="ShapedLightText"/> containing <see cref="GlyphWidth"/> structs
        /// (8 bytes each) plus per-font UnitsPerEm for correct multi-font width calculation.
        /// </summary>
        ShapedLightText ShapeLight(string text, ShapingOptions options = null);

        /// <summary>
        /// Shapes multiple lines (splits on CR/LF/CRLF and shapes each line).
        /// Returns array of shaped lines in font design units.
        /// </summary>
        ShapedText[] ShapeLines(string text, ShapingOptions options = null);

        /// <summary>
        /// Shapes text for vertical layout (top-to-bottom glyph stacking).
        /// Used for Excel vertical text mode (text rotation value 255), where glyphs
        /// are rendered upright and stacked vertically rather than laid out horizontally.
        /// Advance heights are sourced from the 'vmtx' table when available,
        /// with fallback to 'hmtx' advance widths for fonts without vertical metrics.
        /// </summary>
        /// <param name="text">Text to shape</param>
        /// <param name="options">Shaping options</param>
        /// <returns>Shaped vertical text with positioned glyphs in font design units</returns>
        ShapedVerticalText ShapeVertical(string text, ShapingOptions options = null);

        /// <summary>
        /// Shapes text for vertical layout into lightweight <see cref="VerticalGlyphHeight"/> 
        /// structs optimized for vertical text measurement.
        /// Analogous to <see cref="ShapeLight"/> for horizontal text.
        /// Advance heights are sourced from the 'vmtx' table when available,
        /// with fallback to 'hmtx' advance widths for fonts without vertical metrics.
        /// </summary>
        /// <param name="text">Text to shape</param>
        /// <param name="options">Shaping options</param>
        /// <returns>Array of lightweight vertical glyph height structs (8 bytes each)</returns>
        VerticalGlyphHeight[] ShapeLightVertical(string text, ShapingOptions options = null);

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
        /// Gets Ascent (top of container to the baseline) in points
        /// </summary>
        float GetAscentInPoints(float fontSize);

        /// <summary>
        /// Gets descent distance (below baseline) in points.
        /// </summary>
        float GetDescentInPoints(float fontSize);

        /// <summary>
        /// Gets the font's units per em (for manual conversions).
        /// </summary>
        ushort UnitsPerEm { get; }

        /// <summary>
        /// True when this shaper is backed by a real font file and its shaped output carries
        /// usable glyph ids, glyph outlines and font references.
        ///
        /// False for metrics-only shapers, whose output is valid for measurement and line
        /// breaking but carries no glyph identity. Consumers that subset, embed or otherwise
        /// resolve glyphs must check this and refuse rather than emit glyph ids that do not
        /// correspond to any font.
        /// </summary>
        bool HasGlyphIds { get; }
    }
}