namespace FontLab1.Tables.Kern
{
    internal class KernSubTableFormat0
    {
        internal KernSubTableFormat0(MyBinaryReader reader)
        {
            nPairs = reader.ReadUInt16BigEndian();
            var searchRange = reader.ReadUInt16BigEndian();
            var entrySelector = reader.ReadUInt16BigEndian();
            var rangeShift = reader.ReadUInt16BigEndian();
        }

        internal ushort nPairs { get; set; }

        internal KerningPair[] Pairs { get; set; }
    }
}
