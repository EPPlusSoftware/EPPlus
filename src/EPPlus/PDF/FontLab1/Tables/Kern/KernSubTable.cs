namespace FontLab1.Tables.Kern
{
    internal class KernSubTable
    {
        public ushort version { get; set; }

        public ushort length { get; set; }

        public KernCoverage coverage { get; set; }

        public KernSubTableFormat0 Format0Subtable { get; set; }
    }
}
