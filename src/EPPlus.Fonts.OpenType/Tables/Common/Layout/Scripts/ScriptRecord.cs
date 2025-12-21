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

namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.Scripts
{
    /// <summary>
    /// Represents a Script Record within the Script List table.
    /// Each record associates a unique 4-byte script tag with an offset to a Script Table.
    /// </summary>
    public class ScriptRecord : FontTableElement
    {
        /// <summary>
        /// Gets or sets the 4-byte identifier for the script (e.g., 'latn' for Latin).
        /// </summary>
        public Tag ScriptTag { get; set; }

        /// <summary>
        /// Gets or sets the offset to the Script Table, relative to the start of the ScriptList table.
        /// </summary>
        public ushort ScriptOffset { get; set; }

        /// <summary>
        /// Gets or sets the actual Script Table associated with this record.
        /// </summary>
        public ScriptTable ScriptTable { get; set; }

        /// <summary>
        /// Serializes the Script Record.
        /// Note: In most implementations, the ScriptListTable handles the serialization of 
        /// these records to manage the offset backfilling correctly.
        /// </summary>
        /// <param name="writer">The binary writer.</param>
        internal override void Serialize(FontsBinaryWriter writer)
        {
            // The ScriptListTable typically handles this to manage relative offsets.
            // If called directly, it writes the tag and the current ScriptOffset value.
            writer.Write(this.ScriptTag.ToBytes());
            writer.WriteUInt16BigEndian(this.ScriptOffset);
        }
    }
}
