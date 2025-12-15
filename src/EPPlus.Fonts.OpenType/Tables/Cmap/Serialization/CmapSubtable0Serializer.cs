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
using System;

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
