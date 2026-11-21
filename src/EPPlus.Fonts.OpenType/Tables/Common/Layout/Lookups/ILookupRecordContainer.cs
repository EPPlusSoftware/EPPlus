using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups
{
    /// <summary>
    /// Implemented by contextual lookup subtables (GSUB types 5/6, GPOS types 7/8) that
    /// reference other lookups by index through lookup records.
    /// </summary>
    internal interface ILookupRecordContainer
    {
        /// <summary>
        /// Remaps every lookup record's LookupListIndex from the original LookupList to the
        /// subset LookupList. Records whose target lookup was removed are dropped: a lookup is
        /// only removed when it became empty in the subset, so such a record was a no-op and
        /// dropping it preserves the font's shaping behavior.
        /// </summary>
        void RemapLookupListIndices(Dictionary<int, int> oldToNewIndexMap);
    }
}