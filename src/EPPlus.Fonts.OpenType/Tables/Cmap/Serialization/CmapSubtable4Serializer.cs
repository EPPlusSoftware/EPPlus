using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Serialization
{
    internal class CmapSubtable4Serializer
    {
        public void Serialize(CmapSubtable4 table, FontsBinaryWriter writer)
        {
            // Write header fields
            writer.WriteUInt16BigEndian(table.Format);
            writer.WriteUInt16BigEndian((ushort)table.Length);
            writer.WriteUInt16BigEndian((ushort)table.Language);
            writer.WriteUInt16BigEndian(table.SegCountX2);
            writer.WriteUInt16BigEndian(table.SearchRange);
            writer.WriteUInt16BigEndian(table.EntrySelector);
            writer.WriteUInt16BigEndian(table.RangeShift);

            int segCount = table.SegCountX2 / 2;

            // Write segment arrays
            for (int i = 0; i < segCount; i++)
                writer.WriteUInt16BigEndian(table.EndCode[i]);

            writer.WriteUInt16BigEndian(table.ReservedPad);

            for (int i = 0; i < segCount; i++)
                writer.WriteUInt16BigEndian(table.StartCode[i]);

            for (int i = 0; i < segCount; i++)
                writer.WriteInt16BigEndian(table.IdDelta[i]);

            for (int i = 0; i < segCount; i++)
                writer.WriteUInt16BigEndian(table.IdRangeOffset[i]);

            // Write glyphIdArray
            for (int i = 0; i < table.GlyphIdArray.Length; i++)
                writer.WriteUInt16BigEndian(table.GlyphIdArray[i]);
        }
    }
}
