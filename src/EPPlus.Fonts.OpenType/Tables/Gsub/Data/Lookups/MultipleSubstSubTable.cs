/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/09/2026         EPPlus Software AB           GSUB Multiple Substitution (Type 2) support
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    /// <summary>
    /// Represents a Multiple Substitution Subtable (Lookup Type 2, format 1 - the only format
    /// defined for this lookup type).
    /// This lookup replaces a single glyph with a SEQUENCE of one or more glyphs - the opposite
    /// direction of Ligature Substitution (Lookup Type 4), which replaces several glyphs with
    /// one. Fonts commonly use it for the "ccmp" feature to decompose a precomposed glyph into a
    /// base glyph followed by combining mark glyphs.
    /// </summary>
    public class MultipleSubstSubTable : FontTableElement
    {
        /// <summary>
        /// Gets or sets the format identifier. Always 1 - Multiple Substitution has only one
        /// defined subtable format.
        /// </summary>
        public ushort SubtableFormat { get; set; } = 1;

        /// <summary>
        /// Gets or sets the Coverage table which defines the input glyphs to be substituted.
        /// </summary>
        public CoverageTable Coverage { get; set; }

        /// <summary>
        /// Gets or sets the substitute glyph sequences, one per glyph in <see cref="Coverage"/>,
        /// in the same order (i.e. <c>Sequences[Coverage.GetGlyphIndex(gid)]</c> is the
        /// replacement for <c>gid</c>).
        /// </summary>
        public List<ushort[]> Sequences { get; set; } = new List<ushort[]>();

        /// <summary>
        /// Returns the substitute glyph sequence for a given base glyph ID, or null if the glyph
        /// is not covered by this subtable.
        /// </summary>
        /// <param name="baseGlyphId">The original glyph ID.</param>
        /// <returns>The replacement sequence, or null if not covered.</returns>
        public ushort[] GetSubstitution(ushort baseGlyphId)
        {
            int index = Coverage?.GetGlyphIndex(baseGlyphId) ?? -1;

            if (index < 0 || Sequences == null || index >= Sequences.Count)
                return null;

            return Sequences[index];
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            long subTableStart = writer.BaseStream.Position;

            // 1. Write SubstFormat (always 1)
            writer.WriteUInt16BigEndian(1);

            // 2. Placeholder for CoverageOffset (2 bytes)
            long covOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0);

            // 3. SequenceCount and placeholders for SequenceOffsets
            ushort sequenceCount = Sequences != null ? (ushort)Sequences.Count : (ushort)0;
            writer.WriteUInt16BigEndian(sequenceCount);

            long seqOffsetArrayStart = writer.BaseStream.Position;
            for (int i = 0; i < sequenceCount; i++)
            {
                writer.WriteUInt16BigEndian(0);
            }

            // 4. Write each Sequence table, backfilling its offset in the array above
            for (int i = 0; i < sequenceCount; i++)
            {
                long seqOffsetSlot = seqOffsetArrayStart + (i * 2);
                this.WriteRelativeOffset(writer, subTableStart, seqOffsetSlot);

                ushort[] sequence = Sequences[i] ?? new ushort[0];
                writer.WriteUInt16BigEndian((ushort)sequence.Length);
                foreach (ushort gid in sequence)
                {
                    writer.WriteUInt16BigEndian(gid);
                }
            }

            // 5. Serialize CoverageTable and backfill its offset
            if (this.Coverage != null)
            {
                this.WriteRelativeOffset(writer, subTableStart, covOffsetPos);
                this.Coverage.Serialize(writer);
            }
        }
    }
}