using EPPlus.Fonts.OpenType.Tables.Cmap.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    internal class CmapSubtable4_2 : CmapSubtableBase
    {

        public override ushort Format { get; } = 4;

        public override ushort Length { get; internal set; }

        public override  ushort Language { get; internal set; }

        public ushort SegCountX2 { get; internal set; }
        public ushort SearchRange { get; internal set; }
        public ushort EntrySelector { get; internal set; }
        public ushort RangeShift { get; internal set; }

        public ushort[] EndCode { get; internal set; } = new ushort[0];
        public ushort ReservedPad { get; internal set; }
        public ushort[] StartCode { get; internal set; } = new ushort[0];
        public short[] IdDelta { get; internal set; } = new short[0];
        public ushort[] IdRangeOffset { get; internal set; } = new ushort[0];

        public ushort[] GlyphIdArray { get; internal set; } = new ushort[0];

        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable42Serializer();
            serializer.Serialize(this, writer);
        }
    }
}
