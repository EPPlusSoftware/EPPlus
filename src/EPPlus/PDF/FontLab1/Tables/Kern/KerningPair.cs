using System.Diagnostics;

namespace FontLab1.Tables.Kern
{
    [DebuggerDisplay("l: {left}, r: {right}, v: {value}")]
    internal class KerningPair
    {
        public KerningPair(MyBinaryReader reader)
        {
            left = reader.ReadUInt16BigEndian();
            right = reader.ReadUInt16BigEndian();
            value = reader.ReadInt16BigEndian();
        }

        public ushort left { get; set; }

        public ushort right { get; set; }

        public short value { get; set; }
    }
}
