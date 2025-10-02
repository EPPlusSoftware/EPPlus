using System;
using System.Diagnostics;

namespace OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics.TrueTypeFontReader.Tables.Cmap
{
    [DebuggerDisplay("{CharacterCode} - '{Char}': {GlyphIndex}")]
    internal class GlyphMapping
    {
        public ushort CharacterCode { get; set; }

        public ushort GlyphIndex { get; set; }

        public char Char => Convert.ToChar(CharacterCode);

        public override string ToString()
        {
            return Char.ToString() + ": " + GlyphIndex;
        }
    }
}
