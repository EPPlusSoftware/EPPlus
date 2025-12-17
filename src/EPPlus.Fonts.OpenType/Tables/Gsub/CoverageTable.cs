using EPPlus.Fonts.OpenType.Tables.Gsub.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public abstract class CoverageTable : FontTableElement
    {
        public ushort CoverageFormat { get; set; }
        public abstract ushort[] CoveredGlyphs { get; }

        public abstract int GetGlyphIndex(ushort glyphId);

        public abstract ushort[] GetCoveredGlyphs();
    }
}
