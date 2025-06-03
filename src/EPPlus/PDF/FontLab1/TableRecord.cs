namespace FontLab1
{
    internal class TableRecord
    {
        public Tag Tag { get; set; }

        public uint Checksum { get; set; }

        public uint Offset { get; set; }

        public uint Length { get; set; }
    }
}
