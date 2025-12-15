using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class CoverageTableFormat1 : CoverageTable
    {
        public ushort GlyphCount { get; set; }
        public ushort[] GlyphArray { get; set; }
        public override ushort[] CoveredGlyphs => GlyphArray;
    }
}
