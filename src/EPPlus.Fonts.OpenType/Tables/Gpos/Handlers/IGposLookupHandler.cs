/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS lookup handler interface
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Handlers
{
    /// <summary>
    /// Interface for GPOS lookup handlers.
    /// Each lookup type (SinglePos, PairPos, MarkToBase, etc.) implements this interface.
    /// </summary>
    internal interface IGposLookupHandler
    {
        /// <summary>
        /// The GPOS lookup type this handler supports (1-9).
        /// </summary>
        ushort LookupType { get; }

        /// <summary>
        /// Phase 1: Discover additional glyphs that should be included in the subset.
        /// For GPOS, this typically does nothing (positioning doesn't add glyphs).
        /// </summary>
        /// <param name="context">Subsetting context</param>
        /// <param name="lookup">The lookup to analyze</param>
        /// <param name="processor">The GPOS processor (for recursive calls if needed)</param>
        void Discover(FontSubsettingContext context, LookupTable lookup, GposSubsetProcessor processor);

        /// <summary>
        /// Phase 2: Rewrite the lookup with subsetted data and remapped glyph IDs.
        /// </summary>
        /// <param name="context">Subsetting context</param>
        /// <param name="lookup">The original lookup</param>
        /// <returns>Rewritten lookup, or null if no positioning data remains</returns>
        LookupTable Rewrite(FontSubsettingContext context, LookupTable lookup);
    }
}