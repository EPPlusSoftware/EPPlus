namespace FontLab1.Tables.Cmap
{
    internal class CmapSubtable6
    {
        public CmapSubtable6(MyBinaryReader reader)
        {
            _reader = reader;
            var length = _reader.ReadUInt16BigEndian();
            var language = _reader.ReadUInt16BigEndian();
            var firstCode = _reader.ReadUInt16BigEndian();
            var entryCount = _reader.ReadUInt16BigEndian();
            GlyphMappingArray = new GlyphMapping[entryCount];
            for(var x = 0; x < entryCount; x++)
            {
                GlyphMappingArray[x] = new GlyphMapping
                {
                    CharacterCode = (char)(firstCode + x),
                    GlyphIndex = _reader.ReadUInt16BigEndian()
                };
            }
        }

        private readonly MyBinaryReader _reader;

        public GlyphMapping[] GlyphMappingArray { get; set; }
    }
}
