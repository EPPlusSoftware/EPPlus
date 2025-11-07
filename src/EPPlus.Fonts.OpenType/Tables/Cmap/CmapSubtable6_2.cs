using EPPlus.Fonts.OpenType.Tables.Cmap.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    internal class CmapSubtable6_2 : CmapSubtableBase
    {
        public override ushort Format => 6;

        public override ushort Length { get; internal set; }

        public override ushort Language { get; internal set; }

        public ushort FirstCode { get; internal set; }

        public ushort EntryCount { get; internal set; }

        public ushort[] GlyphIdArray { get; internal set; } = new ushort[0];

        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable6_2Serializer();
            serializer.Serialize(this, writer);
        }
    }
}
