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
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    /// <summary>
    /// Represents a Lookup table within the GSUB or GPOS table.
    /// A lookup contains one or more subtables that perform the actual glyph substitutions or positioning.
    /// </summary>
    public class LookupTable : FontTableElement
    {
        /// <summary>
        /// Gets or sets the lookup type (e.g., 1 for Single Substitution, 4 for Ligature Substitution).
        /// </summary>
        public ushort LookupType { get; set; }

        /// <summary>
        /// Gets or sets the lookup qualifiers (e.g., IgnoreBaseGlyphs, IgnoreLigatures).
        /// </summary>
        public ushort LookupFlag { get; set; }

        /// <summary>
        /// Gets or sets the number of subtables contained in this lookup.
        /// </summary>
        public ushort SubTableCount { get; set; }

        /// <summary>
        /// Gets or sets the list of subtables.
        /// </summary>
        public List<FontTableElement> SubTables { get; set; } = new List<FontTableElement>();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // Store the start of the LookupTable to calculate relative offsets for subtables
            long lookupTableStartOffset = writer.BaseStream.Position;

            // 1. Write LookupType
            writer.WriteUInt16BigEndian(this.LookupType);

            // 2. Write LookupFlag
            writer.WriteUInt16BigEndian(this.LookupFlag);

            // 3. Write SubTableCount
            writer.WriteUInt16BigEndian((ushort)this.SubTables.Count);

            // 4. Write placeholders for SubTableOffsets (relative to lookupTableStartOffset)
            List<long> subTableOffsetPositions = new List<long>();
            for (int i = 0; i < this.SubTables.Count; i++)
            {
                subTableOffsetPositions.Add(writer.BaseStream.Position);
                writer.WriteUInt16BigEndian(0);
            }

            // --- Serialize Subtables ---

            for (int i = 0; i < this.SubTables.Count; i++)
            {
                FontTableElement subTable = this.SubTables[i];
                long currentPos = writer.BaseStream.Position;

                // Calculate the relative offset
                ushort relativeSubTableOffset = (ushort)(currentPos - lookupTableStartOffset);

                // Backfill the offset in the offset array
                long subTableOffsetPos = subTableOffsetPositions[i];
                writer.BaseStream.Seek(subTableOffsetPos, SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(relativeSubTableOffset);

                // Return to the current position to serialize the subtable data
                writer.BaseStream.Seek(currentPos, SeekOrigin.Begin);

                // Subtables are responsible for their own internal serialization logic
                switch (this.LookupType)
                {
                    case 1: // Single Substitution
                    case 4: // Ligature Substitution
                    case 6: // Chaining Contextual Substitution
                        subTable.Serialize(writer);
                        break;

                    default:
                        // For other types, we attempt serialization but they may throw NotImplementedException
                        subTable.Serialize(writer);
                        break;
                }
            }
        }
    }
}