using System;
using System.Collections.Generic;

namespace FontLab1.Tables.Cmap
{
    internal class CmapSubtable0
    {
        public CmapSubtable0(MyBinaryReader reader)
        {
            _reader = reader;
            Format = 0;
            Length = _reader.ReadUInt16BigEndian();
            Language = _reader.ReadUInt16BigEndian();
            var mappings = new List<GlyphMapping>();
            for(var c = 0; c < 256; c++)
            {
                var b = reader.ReadByte();
                var ix = BitConverter.ToUInt16(new byte[] { b, 0 }, 0);
                if(ix != 0)
                {
                    mappings.Add(new GlyphMapping
                    {
                        CharacterCode = Convert.ToChar(c),
                        GlyphIndex = ix
                    });
                }
            }
            GlyphMappingArray = mappings.ToArray();
        }

        private readonly MyBinaryReader _reader;

        public ushort Format { get; set; }

        public ushort Length { get; set; }

        public ushort Language { get; set; }

        public GlyphMapping[] GlyphMappingArray { get; set; }
    }
}
