using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Post
{
    internal class PostTableConstants
    {

        // Allowed versions (raw 16.16 fixed values)
        public const int Version10 = 0x00010000; // 1.0
        public const int Version20 = 0x00020000; // 2.0
        public const int Version25 = 0x00025000; // 2.5
        public const int Version30 = 0x00030000; // 3.0

        // Standard Mac glyph name list count for format 2.0 baseline
        public const int StandardMacGlyphNameCount = 258;

        // Reasonable italic angle bounds (soft checks)
        public const int ItalicAngleMaxAbsDegreesWarning = 90; // warn if exceeds this

        // Underline thickness: should be > 0, optionally compare to UnitsPerEm
        public const int MinUnderlineThickness = 1;

    }
}
