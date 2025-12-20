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
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data
{
    /// <summary>
    /// Represents a Feature Table which defines a specific font feature (e.g., ligatures) 
    /// by pointing to one or more lookups in the global Lookup List.
    /// </summary>
    public class FeatureTable : FontTableElement
    {
        /// <summary>
        /// Gets or sets the offset to a feature parameters table. 
        /// Usually set to 0 unless the feature requires specific parameters (e.g., 'size').
        /// </summary>
        public ushort FeatureParams { get; set; }

        /// <summary>
        /// Gets or sets the number of lookups associated with this feature.
        /// </summary>
        public ushort LookupCount { get; set; }

        /// <summary>
        /// Gets or sets an array of indices into the global LookupListTable. 
        /// These indices define which lookups are triggered when this feature is enabled.
        /// </summary>
        public ushort[] LookupListIndices { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // 1. FeatureParams: Offset to parameters (Set to 0 per spec for most features)
            writer.WriteUInt16BigEndian(this.FeatureParams);

            // 2. LookupCount: Number of lookups for this feature
            if (this.LookupListIndices == null)
            {
                writer.WriteUInt16BigEndian(0);
            }
            else
            {
                writer.WriteUInt16BigEndian((ushort)this.LookupListIndices.Length);
            }

            // 3. LookupListIndices: Array of indices into the LookupList
            if (this.LookupListIndices != null)
            {
                foreach (ushort lookupIndex in this.LookupListIndices)
                {
                    writer.WriteUInt16BigEndian(lookupIndex);
                }
            }
        }

        internal FeatureTable Rewrite(FontSubsettingContext context, Dictionary<int, int> lookupMap)
        {
            var newIndices = new List<ushort>();

            if (this.LookupListIndices != null)
            {
                foreach (var oldIndex in this.LookupListIndices)
                {
                    // Kolla om den gamla lookupen finns kvar i vår nya, filtrerade lista
                    if (lookupMap.TryGetValue(oldIndex, out int newIndex))
                    {
                        newIndices.Add((ushort)newIndex);
                    }
                }
            }

            // Om inga lookups finns kvar för denna feature, returnera null 
            // så att Recorden kan rensas bort.
            if (newIndices.Count == 0) return null;

            return new FeatureTable
            {
                FeatureParams = this.FeatureParams,
                LookupCount = (ushort)newIndices.Count,
                LookupListIndices = newIndices.ToArray()
            };
        }
    }
}