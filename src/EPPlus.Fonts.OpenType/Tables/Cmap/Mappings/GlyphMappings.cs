using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Mappings
{

    public class GlyphMappings
    {
        public Dictionary<uint, ushort> CharCodeToGlyphIndex { get; } = new();
        public Dictionary<ushort, List<uint>> GlyphIndexToCharCodes { get; } = new();

        public void AddMapping(uint charCode, ushort glyphIndex)
        {
            CharCodeToGlyphIndex[charCode] = glyphIndex;

            if (!GlyphIndexToCharCodes.TryGetValue(glyphIndex, out var list))
            {
                list = new List<uint>();
                GlyphIndexToCharCodes[glyphIndex] = list;
            }
            list.Add(charCode);
        }

        public ushort? GetGlyphIndex(uint charCode)
            => CharCodeToGlyphIndex.TryGetValue(charCode, out var glyphIndex) ? glyphIndex : null;

        public IEnumerable<uint> GetCharCodes(ushort glyphIndex)
            => GlyphIndexToCharCodes.TryGetValue(glyphIndex, out var codes) ? codes : Enumerable.Empty<uint>();
    }

}
