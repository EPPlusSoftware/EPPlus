namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Hmtx
{
    internal class LongHorMetric
    {
        /// <summary>
        /// Advance width, in font design units.
        /// </summary>
        public ushort advanceWidth{ get; set; }

        /// <summary>
        /// Glyph left side bearing, in font design units.
        /// </summary>
        public short lsb { get; set; }

        public override string ToString()
        {
            return $"advanceWidth: {advanceWidth}, lsb: {lsb}";
        }
    }
}
