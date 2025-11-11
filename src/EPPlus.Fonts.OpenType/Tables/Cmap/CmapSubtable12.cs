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

        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable12Serializer();
            serializer.Serialize(this, writer);
        }
    }
}
