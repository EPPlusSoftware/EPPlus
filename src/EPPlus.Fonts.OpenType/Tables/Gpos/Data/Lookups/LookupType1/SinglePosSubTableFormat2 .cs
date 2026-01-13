/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS Lookup Type 1 Format 2 (Single Adjustment)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType1
{
    /// <summary>
    /// GPOS Lookup Type 1, Format 2: Single Adjustment Positioning
    /// Applies different ValueRecords to each glyph in the Coverage table.
    /// Used when each glyph needs individual positioning (e.g., optical adjustments).
    /// </summary>
    public class SinglePosSubTableFormat2 : GposSubTableBase
    {
        /// <summary>
        /// Format identifier: 2
        /// </summary>
        public ushort SubtableFormat { get; internal set; }

        /// <summary>
        /// Offset to Coverage table (from beginning of SinglePos subtable)
        /// </summary>
        public ushort CoverageOffset { get; internal set; }

        /// <summary>
        /// Defines the types of data in the ValueRecords
        /// </summary>
        public ushort ValueFormat { get; internal set; }

        /// <summary>
        /// Number of ValueRecords (must equal coverage count)
        /// </summary>
        public ushort ValueCount { get; internal set; }

        /// <summary>
        /// Coverage table - defines which glyphs this adjustment applies to
        /// </summary>
        public CoverageTable Coverage { get; internal set; }

        /// <summary>
        /// Array of ValueRecords, one per covered glyph.
        /// Index corresponds to coverage index.
        /// </summary>
        public ValueRecord[] Values { get; internal set; }

        /// <summary>
        /// Tries to get the positioning adjustment for a glyph.
        /// </summary>
        /// <param name="glyphId">The glyph ID to look up</param>
        /// <param name="value">The ValueRecord if found</param>
        /// <returns>True if the glyph is in coverage</returns>
        public bool TryGetAdjustment(ushort glyphId, out ValueRecord value)
        {
            if (Coverage != null)
            {
                int coverageIndex = Coverage.GetGlyphIndex(glyphId);
                if (coverageIndex >= 0 && coverageIndex < Values.Length)
                {
                    value = Values[coverageIndex];
                    return true;
                }
            }

            value = null;
            return false;
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            long subtableStart = writer.BaseStream.Position;

            // Write header
            writer.WriteUInt16BigEndian(SubtableFormat); // 2

            long coverageOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // Coverage offset placeholder

            writer.WriteUInt16BigEndian(ValueFormat);
            writer.WriteUInt16BigEndian(ValueCount);

            // Write ValueRecords array
            if (Values != null)
            {
                foreach (var value in Values)
                {
                    ValueRecordSerializer.Serialize(writer, value, ValueFormat);
                }
            }

            // Write Coverage table
            if (Coverage != null)
            {
                ushort coverageOffset = (ushort)(writer.BaseStream.Position - subtableStart);
                long resumePos = writer.BaseStream.Position;

                // Update offset
                writer.BaseStream.Seek(coverageOffsetPos, System.IO.SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(coverageOffset);
                writer.BaseStream.Seek(resumePos, System.IO.SeekOrigin.Begin);

                // Serialize coverage
                Coverage.Serialize(writer);
            }
        }
    }
}