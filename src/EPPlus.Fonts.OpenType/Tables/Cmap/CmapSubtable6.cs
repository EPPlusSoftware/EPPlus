using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using EPPlus.Fonts.OpenType.Tables.Cmap.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    internal class CmapSubtable6 : CmapSubtableBase
    {
        public override ushort Format => 6;

        public override uint Length { get; internal set; }

        public override uint Language { get; internal set; }

        public ushort FirstCode { get; internal set; }

        public ushort EntryCount { get; internal set; }

        public ushort[] GlyphIdArray { get; internal set; } = new ushort[0];

        public override GlyphMappings GetGlyphMappings()
        {
            var mapping = new GlyphMappings();

            for (int i = 0; i < EntryCount && i < GlyphIdArray.Length; i++)
            {
                uint charCode = (uint)(FirstCode + i);
                ushort glyphIndex = GlyphIdArray[i];

                mapping.AddMapping(charCode, glyphIndex);
            }

            return mapping;
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable6_2Serializer();
            serializer.Serialize(this, writer);
        }
    }
}
