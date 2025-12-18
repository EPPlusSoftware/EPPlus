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
using EPPlus.Fonts.OpenType.Subsetting;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    /// <summary>
    /// Represents a Lookup table in the GSUB table.
    /// A lookup contains one or more subtables of the same type.
    /// </summary>
    public class LookupTable : FontTableElement
    {
        /// <summary>
        /// The type of information this lookup provides (e.g., 1 for Single, 4 for Ligature).
        /// </summary>
        public ushort LookupType { get; set; }

        /// <summary>
        /// Flags that specify how to process the lookup (e.g., IgnoreMarks, RightToLeft).
        /// </summary>
        public ushort LookupFlag { get; set; }

        /// <summary>
        /// Gets or sets the number of subtables. 
        /// Note: When serializing or rewriting, SubTables.Count is used.
        /// </summary>
        public ushort SubTableCount { get; set; }

        /// <summary>
        /// A list of subtables containing the actual substitution data.
        /// </summary>
        public List<FontTableElement> SubTables { get; set; } = new List<FontTableElement>();

        /// <summary>
        /// Optional MarkFilteringSet index in the GDEF table.
        /// </summary>
        public ushort MarkFilteringSet { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            long lookupStart = writer.BaseStream.Position;

            // 1. Write LookupType
            writer.WriteUInt16BigEndian(LookupType);

            // 2. Write LookupFlag
            writer.WriteUInt16BigEndian(LookupFlag);

            // 3. Write SubTableCount
            writer.WriteUInt16BigEndian((ushort)SubTables.Count);

            // 4. Placeholders for SubTable offsets
            long offsetArrayStart = writer.BaseStream.Position;
            for (int i = 0; i < SubTables.Count; i++)
            {
                writer.WriteUInt16BigEndian(0);
            }

            // 5. If UseMarkFilteringSet flag is set (0x0010), write MarkFilteringSet
            if ((LookupFlag & 0x0010) != 0)
            {
                writer.WriteUInt16BigEndian(MarkFilteringSet);
            }

            // --- Serialize SubTables and backfill offsets ---
            for (int i = 0; i < SubTables.Count; i++)
            {
                long currentPos = writer.BaseStream.Position;
                long offsetInArray = offsetArrayStart + (i * 2);

                // Update the offset in the header
                this.WriteRelativeOffset(writer, lookupStart, offsetInArray);

                // Write the subtable data
                SubTables[i].Serialize(writer);
            }
        }

        /// <summary>
        /// Creates a new LookupTable containing only the substitutions relevant to the subset.
        /// </summary>
        internal LookupTable Rewrite(FontSubsettingContext context)
        {
            var newLookup = new LookupTable
            {
                LookupType = this.LookupType,
                LookupFlag = this.LookupFlag,
                MarkFilteringSet = this.MarkFilteringSet
            };

            foreach (var subTable in this.SubTables)
            {
                FontTableElement rewrittenSubTable = null;

                // Route to the correct Rewrite method based on subtable type
                if (subTable is SingleSubstSubTable single)
                    rewrittenSubTable = single.Rewrite(context);
                else if (subTable is LigatureSubstSubTable ligature)
                    rewrittenSubTable = ligature.Rewrite(context);
                else if (subTable is ChainingContextualSubstFormat3 contextual)
                    rewrittenSubTable = contextual.Rewrite(context);
                else if (subTable is ExtensionSubstSubTable extension)
                    rewrittenSubTable = extension.Rewrite(context);

                // Only add the subtable if it still contains valid mappings after subsetting
                if (rewrittenSubTable != null)
                {
                    newLookup.SubTables.Add(rewrittenSubTable);
                }
            }

            // If no subtables remain, this entire lookup is redundant for the subset
            if (newLookup.SubTables.Count == 0) return null;

            return newLookup;
        }
    }
}