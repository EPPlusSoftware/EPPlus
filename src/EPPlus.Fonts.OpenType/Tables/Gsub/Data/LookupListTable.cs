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
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data
{
    /// <summary>
    /// Represents the Lookup List table in an OpenType font.
    /// It contains an array of offsets to all lookup tables used in the GSUB or GPOS table.
    /// </summary>
    public class LookupListTable : FontTableElement
    {
        /// <summary>
        /// Gets or sets the list of Lookup tables.
        /// </summary>
        public List<LookupTable> Lookups { get; set; } = new List<LookupTable>();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // The start position of the LookupList, used for relative offset calculations
            long listStartOffset = writer.BaseStream.Position;

            // 1. Write LookupCount
            writer.WriteUInt16BigEndian((ushort)this.Lookups.Count);

            // 2. Write placeholders for LookupOffsets (2 bytes per lookup)
            List<long> offsetPositions = new List<long>();
            for (int i = 0; i < this.Lookups.Count; i++)
            {
                offsetPositions.Add(writer.BaseStream.Position);
                writer.WriteUInt16BigEndian(0);
            }

            // 3. Serialize each LookupTable and backfill the offsets
            for (int i = 0; i < this.Lookups.Count; i++)
            {
                long currentPos = writer.BaseStream.Position;
                ushort relativeOffset = (ushort)(currentPos - listStartOffset);

                // Return to the offset array to fill in the calculated relative offset
                writer.BaseStream.Seek(offsetPositions[i], SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(relativeOffset);

                // Return to the actual writing position and serialize the LookupTable
                writer.BaseStream.Seek(currentPos, SeekOrigin.Begin);
                this.Lookups[i].Serialize(writer);
            }
        }
    }
}