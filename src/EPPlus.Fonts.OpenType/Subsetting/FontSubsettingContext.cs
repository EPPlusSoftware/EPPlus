/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    /// <summary>
    /// Shared context object passed through the subsetting pipeline.
    /// Holds all mutable state needed by processors.
    /// .NET 3.5 compatible.
    /// </summary>
    public class FontSubsettingContext
    {
        // Immutable – set once in constructor
        public OpenTypeFont OriginalFont { get; private set; }
        public OpenTypeFont SubsetFont { get; private set; }

        // Mutable collections – processors fill these
        public HashSet<uint> UsedCodePoints { get; private set; }
        public HashSet<ushort> IncludedGlyphs { get; private set; }
        public Dictionary<ushort, ushort> OldToNewGlyphId { get; private set; }
        public List<ushort> NewToOldGlyphId { get; private set; }

        /// <summary>
        /// Creates a new subsetting context and initializes all collections.
        /// </summary>
        public FontSubsettingContext(OpenTypeFont originalFont, IEnumerable<int> unicodeChars)
        {
            if (originalFont == null) throw new ArgumentNullException("originalFont");

            OriginalFont = originalFont;
            SubsetFont = new OpenTypeFont(originalFont.Format);

            SubsetFont.AddOrReplaceTable(originalFont.HeadTable.Clone());
            if (originalFont.NameTable != null)
                SubsetFont.AddOrReplaceTable(originalFont.NameTable.Clone());


            UsedCodePoints = new HashSet<uint>();
            IncludedGlyphs = new HashSet<ushort>();
            OldToNewGlyphId = new Dictionary<ushort, ushort>();
            NewToOldGlyphId = new List<ushort>();

            // Fill UsedCodePoints (including space)
            foreach (int c in unicodeChars)
            {
                uint cp = (uint)c;
                if (cp <= 0x10FFFF)
                    UsedCodePoints.Add(cp);
            }

            // Always include space (code point 32)
            UsedCodePoints.Add(32);
        }
    }
}
