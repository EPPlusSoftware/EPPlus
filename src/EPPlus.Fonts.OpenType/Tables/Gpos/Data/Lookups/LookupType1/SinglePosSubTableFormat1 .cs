/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS Lookup Type 1 (Single Adjustment)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType1
{
    /// <summary>
    /// GPOS Lookup Type 1, Format 1: Single Adjustment Positioning
    /// Applies the same ValueRecord to all glyphs in the Coverage table.
    /// Used for uniform adjustments like raising all superscripts by the same amount.
    /// </summary>
    public class SinglePosSubTableFormat1 : GposSubTableBase
    {
        /// <summary>
        /// Format identifier: 1
        /// </summary>
        public ushort SubtableFormat { get; internal set; }

        /// <summary>
        /// Offset to Coverage table (from beginning of SinglePos subtable)
        /// </summary>
        public ushort CoverageOffset { get; internal set; }

        /// <summary>
        /// Defines the types of data in the ValueRecord
        /// </summary>
        public ushort ValueFormat { get; internal set; }

        /// <summary>
        /// Coverage table - defines which glyphs this adjustment applies to
        /// </summary>
        public CoverageTable Coverage { get; internal set; }

        /// <summary>
        /// Single ValueRecord applied to all covered glyphs
        /// </summary>
        public ValueRecord Value { get; internal set; }

        /// <summary>
        /// Tries to get the positioning adjustment for a glyph.
        /// </summary>
        /// <param name="glyphId">The glyph ID to look up</param>
        /// <param name="value">The ValueRecord if found</param>
        /// <returns>True if the glyph is in coverage</returns>
        public bool TryGetAdjustment(ushort glyphId, out ValueRecord value)
        {
            if (Coverage != null && Coverage.IsCovered(glyphId))
            {
                value = Value;
                return true;
            }

            value = null;
            return false;
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            long subtableStart = writer.BaseStream.Position;

            // Write header
            writer.WriteUInt16BigEndian(SubtableFormat); // 1

            long coverageOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // Coverage offset placeholder

            writer.WriteUInt16BigEndian(ValueFormat);

            // Write ValueRecord
            ValueRecordSerializer.Serialize(writer, Value, ValueFormat);

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