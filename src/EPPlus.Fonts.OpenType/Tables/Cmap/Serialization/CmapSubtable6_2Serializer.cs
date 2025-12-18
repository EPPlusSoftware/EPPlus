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
