using System.Collections.Generic;

namespace FontLab1.Tables.Maxp
{
    internal class MaxpTableLoader : TableLoader<MaxpTable>
    {
        public MaxpTableLoader(MyBinaryReader reader, Dictionary<string, TableRecord> tables) : base(reader, tables, TableNames.Maxp)
        {
        }

        protected override MaxpTable LoadInternal()
        {
            var pos = _reader.BaseStream.Position;
            var version = _reader.ReadInt32BigEndian();
            var major = (version >> 16);
            var minor = (version & 16);
            var pos2 = _reader.BaseStream.Position;
            var nGlyphs = _reader.ReadUInt16BigEndian();
            return new MaxpTable
            {
                numGlyphs = nGlyphs
            };
        }
    }
}
