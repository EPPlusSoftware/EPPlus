using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public abstract class SingleSubstSubTable : FontTableElement
    {
        public ushort SubtableFormat { get; set; }
        public CoverageTable Coverage { get; set; }

        // Gemensam metod för att få ut substitutionen för en Base Glyph ID.
        // Måste implementeras i varje format.
        public abstract ushort GetSubstitution(ushort baseGlyphId);
    }
}
