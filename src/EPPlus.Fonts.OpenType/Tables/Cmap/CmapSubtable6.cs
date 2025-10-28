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
    public class CmapSubtable6
    {
        internal CmapSubtable6(FontsBinaryReader reader)
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

        private readonly FontsBinaryReader _reader;

        public GlyphMapping[] GlyphMappingArray { get; set; }
    }
}
