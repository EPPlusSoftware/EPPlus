using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    public class FontSubsettingContext
    {
        public OpenTypeFont OriginalFont { get; }
        public OpenTypeFont SubsetFont { get; }
        public HashSet<uint> UsedCodePoints { get; }
        public HashSet<ushort> IncludedGlyphs { get; }
        public Dictionary<ushort, ushort> OldToNewGlyphId { get; }
        public List<ushort> NewToOldGlyphId { get; }
    }
}
