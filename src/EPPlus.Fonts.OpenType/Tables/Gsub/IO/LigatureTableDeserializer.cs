using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

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
