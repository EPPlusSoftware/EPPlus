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
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Handlers
{
    internal interface IGsubLookupHandler
    {
        ushort LookupType { get; }

        // Phase 1: Identify which glyphs are affected and should be included in the subset
        void Discover(FontSubsettingContext context, LookupTable lookup, GsubSubsetProcessor processor);

        // Phase 2: Create a new, filtered table based on the included glyphs
        LookupTable Rewrite(FontSubsettingContext context, LookupTable oldLookup);
    }
}
