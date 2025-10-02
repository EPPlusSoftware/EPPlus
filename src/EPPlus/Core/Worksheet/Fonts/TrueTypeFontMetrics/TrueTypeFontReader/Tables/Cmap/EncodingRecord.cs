using System.Collections.Generic;

namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Cmap
{
    internal class EncodingRecord
    {
        public EncodingRecord(MyBinaryReader reader)
        {
            _reader = reader;
            PlatformId = (Platforms)reader.ReadUInt16BigEndian();
            EncodingId = reader.ReadUInt16BigEndian();
            SubtableOffset = reader.ReadUInt32BigEndian();
        }

        private readonly MyBinaryReader _reader;
        
        /// <summary>
        /// 0 - Unicode
        /// 1 - Macintosh
        /// 2 - ISO (deprecated)
        /// 3 - Windows
        /// 4 - Custom
        /// </summary>
        public Platforms PlatformId { get; private set; }

       
        public ushort EncodingId { get; private set; }

        public uint SubtableOffset { get; set; }

        public GlyphMapping[] Mappings { get; set; }

        public IDictionary<ushort, char> GlyphIndexToCharMappings { get; internal set; }
        public IDictionary<char, ushort> CharMappingsToGlyphIndex { get; internal set; }
    }
}
