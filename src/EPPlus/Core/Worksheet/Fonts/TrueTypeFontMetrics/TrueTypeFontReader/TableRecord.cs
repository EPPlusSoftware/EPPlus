namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader
{
    internal class TableRecord
    {
        public Tag Tag { get; set; }

        public uint Checksum { get; set; }

        public uint Offset { get; set; }

        public uint Length { get; set; }
    }
}
