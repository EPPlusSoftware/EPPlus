using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{

    public class UvsMapping : FontTableElement
    {
        public uint UnicodeValue { get; internal set; }
        public ushort GlyphId { get; internal set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            writer.WriteUInt24BigEndian(UnicodeValue);
            writer.WriteUInt16BigEndian(GlyphId);
        }
    }

}
