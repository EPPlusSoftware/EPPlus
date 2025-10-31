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
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Serializers
{
    internal class CmapSubtable6Serializer : CmapSubtableSerializerBase<CmapSubtable6>
    {
        internal override void Serialize(CmapSubtable6 subTable, FontsBinaryWriter writer)
        {
            // Format 6 header fields:
            // format (2 bytes) = 6
            // length (2 bytes) = 6 + 2 + 2 + 2 * entryCount = 10 + 2 * entryCount
            // language (2 bytes)
            // firstCode (2 bytes)
            // entryCount (2 bytes)
            // glyphIdArray (2 bytes * entryCount)

            if (subTable.GlyphMappingArray == null || subTable.GlyphMappingArray.Length == 0)
            {
                throw new InvalidOperationException("GlyphMappingArray is empty. Cannot serialize CmapSubtable6.");
            }

            // Determine firstCode and entryCount
            ushort firstCode = subTable.GlyphMappingArray.Min(g => g.CharacterCode);
            ushort lastCode = subTable.GlyphMappingArray.Max(g => g.CharacterCode);
            ushort entryCount = (ushort)(lastCode - firstCode + 1);

            // Build glyphIdArray with default value 0
            ushort[] glyphIdArray = new ushort[entryCount];
            foreach (var mapping in subTable.GlyphMappingArray)
            {
                int index = mapping.CharacterCode - firstCode;
                glyphIdArray[index] = mapping.GlyphIndex;
            }

            // Calculate total length
            ushort length = (ushort)(10 + entryCount * 2);

            // Write header
            writer.WriteUInt16BigEndian(subTable.Format);       // format = 6
            writer.WriteUInt16BigEndian(length);       // total length
            writer.WriteUInt16BigEndian(subTable.Language);     // language
            writer.WriteUInt16BigEndian(firstCode);    // first character code
            writer.WriteUInt16BigEndian(entryCount);   // number of entries

            // Write glyphIdArray
            foreach (var glyphIndex in glyphIdArray)
            {
                writer.WriteUInt16BigEndian(glyphIndex);
            }
        }
    }
}
