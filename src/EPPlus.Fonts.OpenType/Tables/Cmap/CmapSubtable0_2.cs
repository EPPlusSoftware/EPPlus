using EPPlus.Fonts.OpenType.Tables.Cmap.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    internal class CmapSubtable0_2 : CmapSubtableBase
    {
        public CmapSubtable0_2()
        {
            GlyphIdArray = new byte[256];
        }

        public override ushort Format { get { return 0; } }

        public override ushort Length { get; internal set; }

        public override ushort Language { get; internal set; }

        /// <summary>
        /// Maps character codes 0–255 to glyph indices.
        /// </summary>
        public byte[] GlyphIdArray { get; internal set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable02Serializer();
            serializer.Serialize(this, writer);

        }
    }
}
