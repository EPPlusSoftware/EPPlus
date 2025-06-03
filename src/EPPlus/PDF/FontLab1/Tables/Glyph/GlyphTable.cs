namespace FontLab1.Tables.Glyph
{
    /// <summary>
    /// This table contains information that describes the glyphs in the font in the TrueType outline format
    /// https://docs.microsoft.com/en-us/typography/opentype/spec/glyf
    /// </summary>
    internal class GlyphTable
    {
        public GlyphHeader[] Glyphs { get; set; }
    }
}
