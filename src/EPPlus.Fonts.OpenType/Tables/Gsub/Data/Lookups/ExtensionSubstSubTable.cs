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
  12/21/2025         EPPlus Software AB           Refactor: Inherit from ExtensionSubTableBase
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Subsetting;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    /// <summary>
    /// Represents an Extension Substitution subtable (Lookup Type 7).
    /// This is used to reference subtables that exceed the 16-bit offset limit.
    /// </summary>
    public class ExtensionSubstSubTable : ExtensionSubTableBase
    {
        /// <summary>
        /// Rewrites the extension by rewriting the inner subtable.
        /// </summary>
        internal ExtensionSubstSubTable Rewrite(FontSubsettingContext context, LookupTable oldLookup)
        {
            if (ExtendedSubTable == null) return null;

            // We delegate the rewrite to the specific type of the inner table
            FontTableElement rewrittenInner = null;

            if (ExtendedSubTable is SingleSubstSubTable single)
                rewrittenInner = single.Rewrite(context, oldLookup);
            else if (ExtendedSubTable is LigatureSubstSubTable ligature)
                rewrittenInner = ligature.Rewrite(context, oldLookup);
            else if (ExtendedSubTable is ChainingContextualSubstFormat3 contextual)
                rewrittenInner = contextual.Rewrite(context);
            // Add more types here as they are implemented

            if (rewrittenInner == null) return null;

            return new ExtensionSubstSubTable
            {
                ExtensionLookupType = this.ExtensionLookupType,
                ExtendedSubTable = rewrittenInner
            };
        }
    }
}