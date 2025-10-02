namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Kern
{
    internal class KernTable
    {
        public ushort version { get; set; }

        public ushort nTables { get; set; }

        public KernSubTable[] SubTables { get; set; }

        public ushort NumberOfFormat0Tables { get; set; }
    }
}
