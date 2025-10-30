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
using System.Linq;
using System;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    public class CmapSubtable6 : CmapSubtableBase
    {
        internal CmapSubtable6(FontsBinaryReader reader)
        {
            _reader = reader;
            Format = 6;
            Length = _reader.ReadUInt16BigEndian();
            Language = _reader.ReadUInt16BigEndian();
            var firstCode = _reader.ReadUInt16BigEndian();
            var entryCount = _reader.ReadUInt16BigEndian();
            GlyphMappingArray = new GlyphMapping[entryCount];
            for(var x = 0; x < entryCount; x++)
            {
                GlyphMappingArray[x] = new GlyphMapping
                {
                    CharacterCode = (char)(firstCode + x),
                    GlyphIndex = _reader.ReadUInt16BigEndian()
                };
            }
        }

        private readonly FontsBinaryReader _reader;

        public override ushort Format { get; }

        public override ushort Length => (ushort)(10 + GlyphMappingArray.Length* 2)
        public override ushort Language { get; }

        public override GlyphMapping[] GlyphMappingArray { get; }

        internal override void Serialize(FontsBinaryWriter writer)
        {

            // Format 6 header fields:
            // format (2 bytes) = 6
            // length (2 bytes) = 6 + 2 + 2 + 2 * entryCount = 10 + 2 * entryCount
            // language (2 bytes)
            // firstCode (2 bytes)
            // entryCount (2 bytes)
            // glyphIdArray (2 bytes * entryCount)

            if (GlyphMappingArray == null || GlyphMappingArray.Length == 0)
            {
                throw new InvalidOperationException("GlyphMappingArray is empty. Cannot serialize CmapSubtable6.");
            }

            // Determine firstCode and entryCount
            ushort firstCode = GlyphMappingArray.Min(g => g.CharacterCode);
            ushort lastCode = GlyphMappingArray.Max(g => g.CharacterCode);
            ushort entryCount = (ushort)(lastCode - firstCode + 1);

            // Build glyphIdArray with default value 0
            ushort[] glyphIdArray = new ushort[entryCount];
            foreach (var mapping in GlyphMappingArray)
            {
                int index = mapping.CharacterCode - firstCode;
                glyphIdArray[index] = mapping.GlyphIndex;
            }

            // Calculate total length
            ushort length = (ushort)(10 + entryCount * 2);

            // Write header
            writer.WriteUInt16BigEndian(Format);       // format = 6
            writer.WriteUInt16BigEndian(length);       // total length
            writer.WriteUInt16BigEndian(Language);     // language
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
