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

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data
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
    }
}