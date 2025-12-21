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
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Common.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    /// <summary>
    /// Represents a Single Substitution Subtable (Lookup Type 1).
    /// This lookup replaces a single glyph with another single glyph.
    /// </summary>
    public abstract class SingleSubstSubTable : FontTableElement
    {
        /// <summary>
        /// Gets or sets the format identifier (1 or 2).
        /// </summary>
        public ushort SubtableFormat { get; set; }

        /// <summary>
        /// Gets or sets the Coverage table which defines the input glyphs to be substituted.
        /// </summary>
        public CoverageTable Coverage { get; set; }

        /// <summary>
        /// Returns the substituted glyph ID for a given base glyph ID.
        /// </summary>
        /// <param name="baseGlyphId">The original glyph ID.</param>
        /// <returns>The new glyph ID after substitution.</returns>
        public abstract ushort GetSubstitution(ushort baseGlyphId);

        /// <summary>
        /// Creates a subset of the subtable based on the provided mapping.
        /// Note: This implementation maps everything to Format 2 to ensure compatibility 
        /// when indices are no longer contiguous.
        /// </summary>
        internal virtual SingleSubstSubTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
        {
            // Listan för att hålla (Nytt Input GID, Nytt Output GID)
            List<GsubRewriteEntry> validMappings = new List<GsubRewriteEntry>();

            // Hämta alla Glyph IDs som denna tabell hanterar från original-fonten
            ushort[] oldInputGlyphs = this.Coverage.GetCoveredGlyphs();

            foreach (ushort oldInputGid in oldInputGlyphs)
            {
                // 1. Finns tecknet som triggar bytet (t.ex. 'f') i vårt subset?
                if (context.OldToNewGlyphId.TryGetValue(oldInputGid, out ushort newInputGid))
                {
                    // Hämta vad tecknet skulle bytas ut mot (t.ex. GID 447)
                    ushort oldOutputGid = GetSubstitution(oldInputGid);

                    // 2. Finns ersättningstecknet (GID 447) också i vårt subset?
                    // DETTA ÄR KRITISKT: Om 447 inte finns i IncludedGlyphs blir det ingen mappning.
                    if (context.OldToNewGlyphId.TryGetValue(oldOutputGid, out ushort newOutputGid))
                    {
                        validMappings.Add(new GsubRewriteEntry
                        {
                            NewInput = newInputGid,
                            NewOutput = newOutputGid
                        });
                    }
                }
            }

            if (validMappings.Count == 0) return null;

            // Sortera efter NewInput - ett strikt krav för CoverageTable i OpenType
            validMappings.Sort((a, b) => a.NewInput.CompareTo(b.NewInput));

            // Skapa den nya tabellen som Format 2 (det säkraste för subsetting)
            var newTable = new SingleSubstSubTableFormat2();
            List<ushort> newInputs = new List<ushort>();
            newTable.SubstituteGlyphIDs = new ushort[validMappings.Count];

            for (int i = 0; i < validMappings.Count; i++)
            {
                newInputs.Add(validMappings[i].NewInput);
                newTable.SubstituteGlyphIDs[i] = validMappings[i].NewOutput;
            }

            // Bygg om Coverage med de NYA GID-numren
            newTable.Coverage = CoverageTableFormat2.CreateCoverageFormat2(newInputs);
            newTable.GlyphCount = (ushort)validMappings.Count;
            newTable.SubtableFormat = 2;

            return newTable;
        }

        private struct GsubRewriteEntry
        {
            public ushort NewInput;
            public ushort NewOutput;
        }
    }
}