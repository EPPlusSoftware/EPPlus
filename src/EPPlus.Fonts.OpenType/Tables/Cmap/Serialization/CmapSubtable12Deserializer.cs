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
namespace EPPlus.Fonts.OpenType.Tables.Cmap.Serialization
{
    internal class CmapSubtable12Deserializer
    {
        private readonly FontsBinaryReader _reader;

        public CmapSubtable12Deserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public CmapSubtable12 Deserialize(uint startIndex)
        {
            _reader.BaseStream.Position = startIndex;

            var subtable = new CmapSubtable12();

            // Read header
            var format = _reader.ReadUInt16BigEndian(); // should always be 12
            var reserved = _reader.ReadUInt16BigEndian(); // always 0
            subtable.Length = _reader.ReadUInt32BigEndian();
            subtable.Language = _reader.ReadUInt32BigEndian();
            subtable.NumGroups = _reader.ReadUInt32BigEndian();

            // Read groups
            for (int i = 0; i < subtable.NumGroups; i++)
            {
                var group = new SequencialMapGroup
                {
                    StartCharCode = _reader.ReadUInt32BigEndian(),
                    EndCharCode = _reader.ReadUInt32BigEndian(),
                    StartGlyphId = _reader.ReadUInt32BigEndian()
                };
                subtable.Groups.Add(group);
            }

            return subtable;
        }
    }
}
