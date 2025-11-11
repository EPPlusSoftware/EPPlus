using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Mappings
{
    internal class CharGlyphPair
    {
        public uint CharCode { get; set; }
        public ushort GlyphIndex { get; set; }

        public CharGlyphPair(uint charCode, ushort glyphIndex)
        {
            CharCode = charCode;
            GlyphIndex = glyphIndex;
        }

    }
}
