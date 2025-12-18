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
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.IO
{
    internal class LigatureTableDeserializer
    {
        private readonly FontsBinaryReader _reader;

        public LigatureTableDeserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public LigatureTable Deserialize(long startIndex)
        {
            _reader.BaseStream.Seek(startIndex, SeekOrigin.Begin);

            LigatureTable ligTable = new LigatureTable();

            // USHORT LigatureGlyph (output)
            ligTable.LigatureGlyph = _reader.ReadUInt16BigEndian();

            // USHORT ComponentCount (number of components following BaseGlyph)
            ushort componentCount = _reader.ReadUInt16BigEndian();

            // USHORT[] ComponentGlyphIDs
            ligTable.Components = new ushort[componentCount];
            for (int i = 0; i < componentCount; i++)
            {
                ligTable.Components[i] = _reader.ReadUInt16BigEndian();
            }

            return ligTable;
        }
    }
}
