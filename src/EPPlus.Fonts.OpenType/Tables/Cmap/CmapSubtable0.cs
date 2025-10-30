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
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    public class CmapSubtable0 : CmapSubtableBase
    {
        internal CmapSubtable0(FontsBinaryReader reader)
        {
            _reader = reader;
            Format = 0;
            Length = _reader.ReadUInt16BigEndian();
            Language = _reader.ReadUInt16BigEndian();
            var mappings = new List<GlyphMapping>();
            for(var c = 0; c < 256; c++)
            {
                var b = reader.ReadByte();
                var ix = BitConverter.ToUInt16(new byte[] { b, 0 }, 0);
                if(ix != 0)
                {
                    mappings.Add(new GlyphMapping
                    {
                        CharacterCode = Convert.ToChar(c),
                        GlyphIndex = ix
                    });
                }
            }
            GlyphMappingArray = mappings.ToArray();
        }

        private readonly FontsBinaryReader _reader;

        public override ushort Format { get; }

        public override ushort Length { get; }

        public override ushort Language { get; }

        public override GlyphMapping[] GlyphMappingArray { get; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // Write the subtable header: format, length, and language
            writer.WriteUInt16BigEndian(Format);   // Format = 0
            writer.WriteUInt16BigEndian(Length);   // Total length of the subtable (should be 262 bytes)
            writer.WriteUInt16BigEndian(Language); // Language code

            // Create a 256-byte array for the glyphIdArray (1 byte per character code 0–255)
            byte[] glyphIdArray = new byte[256];

            foreach (var mapping in GlyphMappingArray)
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
