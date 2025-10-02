namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Loca
{
    /// <summary>
    /// The indexToLoc table stores the offsets to the locations of the glyphs in the font, relative to the beginning of the glyphData table
    /// https://docs.microsoft.com/en-us/typography/opentype/spec/loca
    /// </summary>
    internal class LocaTable
    {
        public uint[] Offsets { get; set; }
    }
}
