using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Serialization
{
    internal class CmapSubtable02Serializer
    {
        public void Serialize(CmapSubtable0 table, FontsBinaryWriter writer)
        {
            // Format 0 requires exactly 256 bytes in the glyphIdArray
            if (table.GlyphIdArray == null || table.GlyphIdArray.Length != 256)
                throw new InvalidOperationException("GlyphIdArray must contain exactly 256 entries for format 0.");

            // Format 0 has a fixed length: 6 bytes header + 256 bytes glyphIdArray = 262 bytes
            if (table.Length == 0)
                table.Length = 262;

            // Write the format (always 0)
            writer.WriteUInt16BigEndian(table.Format);

            // Write the total length of the subtable
            writer.WriteUInt16BigEndian((ushort)table.Length);

            // Write the language field
            writer.WriteUInt16BigEndian((ushort)table.Language);

            // Write the 256-byte glyphIdArray
            writer.Write(table.GlyphIdArray);
        }
    }

}
