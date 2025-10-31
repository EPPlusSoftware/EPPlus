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

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Serializers
{
    internal class CmapSubtable0Serializer : CmapSubtableSerializerBase<CmapSubtable0>
    {
        internal override void Serialize(CmapSubtable0 subTable, FontsBinaryWriter writer)
        {
            // Write the subtable header: format, length, and language
            writer.WriteUInt16BigEndian(subTable.Format);   // Format = 0
            writer.WriteUInt16BigEndian(subTable.Length);   // Total length of the subtable (should be 262 bytes)
            writer.WriteUInt16BigEndian(subTable.Language); // Language code

            // Create a 256-byte array for the glyphIdArray (1 byte per character code 0–255)
            byte[] glyphIdArray = new byte[256];

            foreach (var mapping in subTable.GlyphMappingArray)
            {
                // Format 0 only supports character codes in the range 0–255
                if (mapping.CharacterCode >= 256)
                {
                    throw new InvalidOperationException(
                        $"Character code {mapping.CharacterCode} is out of range for format 0 (must be < 256).");
                }

                // Format 0 only supports glyph indices in the range 0–255 (1 byte)
                if (mapping.GlyphIndex > 255)
                {
                    throw new InvalidOperationException(
                        $"Glyph index {mapping.GlyphIndex} for character code {mapping.CharacterCode} exceeds 255 and cannot be encoded in format 0.");
                }

                glyphIdArray[mapping.CharacterCode] = (byte)mapping.GlyphIndex;
            }

            // Write the glyphIdArray to the stream
            writer.Write(glyphIdArray);
        }
    }
}
