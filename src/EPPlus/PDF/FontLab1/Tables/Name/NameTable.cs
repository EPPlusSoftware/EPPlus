namespace FontLab1.Tables.Name
{
    internal class NameTable
    {
        public ushort format { get; set; }

        public ushort count { get; set; }

        public ushort stringOffset { get; set; }

        public NameRecord[] NameRecords { get; set; }
    }
}
