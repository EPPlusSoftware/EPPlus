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

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data
{
    /// <summary>
    /// Represents a Language System Table, which defines the features available 
    /// for a particular language system within a script.
    /// </summary>
    public class LangSysTable : FontTableElement
    {
        /// <summary>
        /// Gets or sets the lookup order. Reserved for future use; currently set to 0.
        /// </summary>
        public ushort LookupOrder { get; set; }

        /// <summary>
        /// Gets or sets the index of a required feature in the FeatureList. 
        /// Set to 0xFFFF if no required feature is defined.
        /// </summary>
        public ushort RequiredFeatureIndex { get; set; }

        /// <summary>
        /// Gets or sets the number of optional features associated with this language system.
        /// </summary>
        public ushort FeatureIndexCount { get; set; }

        /// <summary>
        /// Gets or sets an array of indices into the FeatureList for the optional features.
        /// </summary>
        public ushort[] FeatureIndices { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // 1. LookupOrder: Reserved, should be 0
            writer.WriteUInt16BigEndian(this.LookupOrder);

            // 2. RequiredFeatureIndex: Index into FeatureList or 0xFFFF
            writer.WriteUInt16BigEndian(this.RequiredFeatureIndex);

            // 3. FeatureIndexCount: Number of optional features
            ushort count = this.FeatureIndices != null ? (ushort)this.FeatureIndices.Length : (ushort)0;
            writer.WriteUInt16BigEndian(count);

            // 4. FeatureIndices: Array of indices into the FeatureList
            if (this.FeatureIndices != null)
            {
                foreach (ushort index in this.FeatureIndices)
                {
                    writer.WriteUInt16BigEndian(index);
                }
            }
        }
    }
}