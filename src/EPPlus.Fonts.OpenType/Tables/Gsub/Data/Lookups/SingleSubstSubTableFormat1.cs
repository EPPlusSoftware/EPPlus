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
    /// Represents a Single Substitution Subtable Format 1.
    /// This format applies a single constant delta value to a range of glyph IDs.
    /// </summary>
    public class SingleSubstSubTableFormat1 : SingleSubstSubTable
    {
        /// <summary>
        /// Gets or sets the delta value added to the original glyph ID to get the substituted glyph ID.
        /// </summary>
        public short DeltaGlyphID { get; set; }

        /// <summary>
        /// Calculates the substituted glyph ID by adding the delta to the input glyph ID.
        /// </summary>
        /// <param name="baseGlyphId">The original glyph ID.</param>
        /// <returns>The substituted glyph ID if covered; otherwise, 0.</returns>
        public override ushort GetSubstitution(ushort baseGlyphId)
        {
            int index = Coverage.GetGlyphIndex(baseGlyphId);
            if (index == -1) return 0; // Glyph is not covered by this subtable

            // Format 1 adds a constant delta to the original GID.
            // Wrap-around (modulo 65536) is handled by the ushort cast.
            return (ushort)(baseGlyphId + DeltaGlyphID);
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // Store the start position of this subtable for relative offset calculation
            long subTableStart = writer.BaseStream.Position;

            // 1. Write SubtableFormat (1)
            writer.WriteUInt16BigEndian(1);

            // 2. Placeholder for CoverageOffset (2 bytes)
            long covOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0);

            // 3. Write DeltaGlyphID (SSHORT)
            writer.WriteInt16BigEndian(this.DeltaGlyphID);

            // --- Write CoverageTable ---
            if (this.Coverage != null)
            {
                // Calculate and backfill the relative offset to the Coverage table
                this.WriteRelativeOffset(writer, subTableStart, covOffsetPos);

                // Serialize the CoverageTable
                this.Coverage.Serialize(writer);
            }
        }
    }
}