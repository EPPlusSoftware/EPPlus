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
