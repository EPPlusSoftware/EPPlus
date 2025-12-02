using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Serialization
{
    internal class CmapSubtable6_2Serializer
    {
        public void Serialize(CmapSubtable6 table, FontsBinaryWriter writer)
        {
            if (table.GlyphIdArray == null || table.GlyphIdArray.Length != table.EntryCount)
                throw new InvalidOperationException("GlyphIdArray length must match EntryCount.");

            // Format 6 length = 10 bytes header + 2 * entryCount
            if (table.Length == 0)
                table.Length = (ushort)(10 + 2 * table.EntryCount);

            writer.WriteUInt16BigEndian(table.Format);
            writer.WriteUInt16BigEndian((ushort)table.Length);
            writer.WriteUInt16BigEndian((ushort)table.Language);
            writer.WriteUInt16BigEndian(table.FirstCode);
            writer.WriteUInt16BigEndian(table.EntryCount);

            for (int i = 0; i < table.EntryCount; i++)
                writer.WriteUInt16BigEndian(table.GlyphIdArray[i]);
        }
    }
}
