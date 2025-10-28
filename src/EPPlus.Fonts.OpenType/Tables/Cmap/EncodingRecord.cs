using System.Collections.Generic;

/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    public class EncodingRecord
    {
        internal EncodingRecord(FontsBinaryReader reader)
        {
            _reader = reader;
            PlatformId = (Platforms)reader.ReadUInt16BigEndian();
            EncodingId = reader.ReadUInt16BigEndian();
            SubtableOffset = reader.ReadUInt32BigEndian();
        }

        private readonly FontsBinaryReader _reader;
        
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
