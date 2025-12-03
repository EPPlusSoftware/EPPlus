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
using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using EPPlus.Fonts.OpenType.Tables.Cmap.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    internal class CmapSubtable12 : CmapSubtableBase
    {
        public override ushort Format { get; } = 12;

        public override uint Length { get; internal set; }

        public override uint Language { get; internal set; }

        public ushort Reserved { get; } = 0;

        public uint NumGroups { get; internal set; }

        public List<SequencialMapGroup> Groups { get; } = new List<SequencialMapGroup>();

        public override GlyphMappings GetGlyphMappings()
        {
            var mapping = new GlyphMappings();

            foreach (var group in Groups)
            {
                uint startCharCode = group.StartCharCode;
                uint endCharCode = group.EndCharCode;
                uint startGlyphId = group.StartGlyphId;

                for (uint charCode = startCharCode; charCode <= endCharCode; charCode++)
                {
                    ushort glyphIndex = (ushort)(startGlyphId + (charCode - startCharCode));
                    mapping.AddMapping(charCode, glyphIndex);
                }
            }

            return mapping;
        }


        internal override int MapCodePointToGlyph(int codePoint)
        {
            int left = 0;
            int right = Groups.Count - 1;

            while (left <= right)
            {
                int mid = (left + right) / 2;
                var group = Groups[mid];

                if (codePoint < group.StartCharCode)
                    right = mid - 1;
                else if (codePoint > group.EndCharCode)
                    left = mid + 1;
                else
                    return (int)(group.StartGlyphId + (codePoint - group.StartCharCode));
            }

            return -1;
        }


        public override bool TryGetGlyphId(int codePoint, out ushort glyphId)
        {
            glyphId = 0;

            foreach (var group in Groups)
            {
                if (codePoint >= group.StartCharCode && codePoint <= group.EndCharCode)
                {
                    uint offset = (uint)(codePoint - group.StartCharCode);
                    glyphId = (ushort)(group.StartGlyphId + offset);
                    return glyphId != 0;
                }
            }

            return false;
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable12Serializer();
            serializer.Serialize(this, writer);
        }
    }
}
