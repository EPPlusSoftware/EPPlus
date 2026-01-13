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
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Features
{
    /// <summary>
    /// Represents a Feature Record within a Feature List Table.
    /// Associates a 4-byte feature tag with a specific Feature Table.
    /// </summary>
    public class FeatureRecord : FontTableElement
    {
        /// <summary>
        /// Gets or sets the 4-byte feature identifier tag (e.g., 'liga', 'kern').
        /// </summary>
        public Tag FeatureTag { get; set; }

        /// <summary>
        /// Gets or sets the offset to the Feature Table, relative to the start of the Feature List Table.
        /// </summary>
        public ushort FeatureOffset { get; set; }

        /// <summary>
        /// Gets or sets the actual Feature Table associated with this record.
        /// </summary>
        public FeatureTable FeatureTable { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            throw new NotImplementedException();
        }

        internal FeatureRecord Rewrite(FontSubsettingContext context, Dictionary<int, int> lookupMap)
        {
            if (this.FeatureTable == null) return null;

            // Skapa en ny tabell baserat på den gamla
            var rewrittenTable = this.FeatureTable.Rewrite(context, lookupMap);

            // Om tabellen inte längre pekar på några lookups, kastar vi hela recorden
            if (rewrittenTable == null || rewrittenTable.LookupListIndices.Length == 0)
            {
                return null;
            }

            return new FeatureRecord
            {
                FeatureTag = this.FeatureTag,
                FeatureTable = rewrittenTable
                // FeatureOffset räknas ut under serialiseringen senare
            };
        }
    }
}