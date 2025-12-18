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
namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    /// <summary>
    /// Represents a Single Substitution Subtable Format 2.
    /// This format maps each glyph identified in the Coverage table to a specific 
    /// substitute glyph ID in a corresponding array.
    /// </summary>
    public class SingleSubstSubTableFormat2 : SingleSubstSubTable
    {
        /// <summary>
        /// Gets or sets the number of glyph IDs in the SubstituteGlyphIDs array.
        /// </summary>
        public ushort GlyphCount { get; set; }

        /// <summary>
        /// Gets or sets the array of substitute glyph IDs, ordered by their corresponding 
        /// index in the Coverage table.
        /// </summary>
        public ushort[] SubstituteGlyphIDs { get; set; }

        /// <summary>
        /// Returns the substituted glyph ID by looking up the coverage index in the SubstituteGlyphIDs array.
        /// </summary>
        /// <param name="baseGlyphId">The original glyph ID.</param>
        /// <returns>The substituted glyph ID if covered; otherwise, 0.</returns>
        public override ushort GetSubstitution(ushort baseGlyphId)
        {
            int index = Coverage.GetGlyphIndex(baseGlyphId);

            // Validate that the glyph is covered and the index is within the bounds of our array
            if (index == -1 || index >= SubstituteGlyphIDs.Length)
                return 0;

            // Format 2 maps the coverage index directly to the substitute array
            return SubstituteGlyphIDs[index];
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // Store start position for relative offset calculation
            long startPos = writer.BaseStream.Position;

            // 1. Write SubtableFormat (2)
            writer.WriteUInt16BigEndian(2);

            // 2. Placeholder for CoverageOffset (2 bytes)
            long covOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0);

            // 3. Write GlyphCount and the array of SubstituteGlyphIDs
            ushort count = this.SubstituteGlyphIDs != null ? (ushort)this.SubstituteGlyphIDs.Length : (ushort)0;
            writer.WriteUInt16BigEndian(count);

            if (this.SubstituteGlyphIDs != null)
            {
                foreach (var gid in this.SubstituteGlyphIDs)
                {
                    writer.WriteUInt16BigEndian(gid);
                }
            }

            // 4. Serialize CoverageTable and backfill the offset
            if (this.Coverage != null)
            {
                this.WriteRelativeOffset(writer, startPos, covOffsetPos);
                this.Coverage.Serialize(writer);
            }
        }
    }
}