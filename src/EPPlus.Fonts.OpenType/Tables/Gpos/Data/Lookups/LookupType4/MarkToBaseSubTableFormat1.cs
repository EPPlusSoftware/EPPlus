/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           GPOS Lookup Type 4 (MarkToBase Attachment)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4
{
    /// <summary>
    /// GPOS Lookup Type 4, Format 1: MarkToBase Attachment Positioning
    /// Positions mark glyphs (accents, diacritics) relative to base glyphs (letters).
    /// Used for proper placement of é, ñ, ü, etc.
    /// </summary>
    public class MarkToBaseSubTableFormat1 : GposSubTableBase
    {
        /// <summary>
        /// Format identifier: 1 (only format for MarkToBase)
        /// </summary>
        public ushort SubtableFormat { get; internal set; }

        /// <summary>
        /// Offset to MarkCoverage table (from beginning of MarkToBase subtable)
        /// </summary>
        public ushort MarkCoverageOffset { get; internal set; }

        /// <summary>
        /// Offset to BaseCoverage table (from beginning of MarkToBase subtable)
        /// </summary>
        public ushort BaseCoverageOffset { get; internal set; }

        /// <summary>
        /// Number of mark classes (different attachment types)
        /// </summary>
        public ushort MarkClassCount { get; internal set; }

        /// <summary>
        /// Offset to MarkArray table
        /// </summary>
        public ushort MarkArrayOffset { get; internal set; }

        /// <summary>
        /// Offset to BaseArray table
        /// </summary>
        public ushort BaseArrayOffset { get; internal set; }

        /// <summary>
        /// Coverage table defining which mark glyphs (accents) are covered
        /// </summary>
        public CoverageTable MarkCoverage { get; internal set; }

        /// <summary>
        /// Coverage table defining which base glyphs (letters) are covered
        /// </summary>
        public CoverageTable BaseCoverage { get; internal set; }

        /// <summary>
        /// Array of mark records (one per mark glyph in coverage order)
        /// </summary>
        public MarkArray MarkArray { get; internal set; }

        /// <summary>
        /// Array of base records (one per base glyph in coverage order)
        /// </summary>
        public BaseArray BaseArray { get; internal set; }

        /// <summary>
        /// Tries to get the positioning for a mark attached to a base.
        /// </summary>
        /// <param name="markGlyphId">The mark glyph ID (e.g., combining acute accent)</param>
        /// <param name="baseGlyphId">The base glyph ID (e.g., letter 'e')</param>
        /// <param name="markAnchor">Anchor point on the mark</param>
        /// <param name="baseAnchor">Anchor point on the base</param>
        /// <returns>True if attachment is defined</returns>
        public bool TryGetAttachment(ushort markGlyphId, ushort baseGlyphId,
            out AnchorTable markAnchor, out AnchorTable baseAnchor)
        {
            markAnchor = null;
            baseAnchor = null;

            // Get mark index and class
            int markIndex = MarkCoverage?.GetGlyphIndex(markGlyphId) ?? -1;
            if (markIndex < 0 || markIndex >= MarkArray?.Records.Length)
                return false;

            var markRecord = MarkArray.Records[markIndex];
            ushort markClass = markRecord.MarkClass;

            // Get base index
            int baseIndex = BaseCoverage?.GetGlyphIndex(baseGlyphId) ?? -1;
            if (baseIndex < 0 || baseIndex >= BaseArray?.Records.Length)
                return false;

            var baseRecord = BaseArray.Records[baseIndex];

            // Get anchors for this mark class
            if (markClass >= baseRecord.BaseAnchors.Length)
                return false;

            markAnchor = markRecord.MarkAnchor;
            baseAnchor = baseRecord.BaseAnchors[markClass];

            return markAnchor != null && baseAnchor != null;
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            long subtableStart = writer.BaseStream.Position;

            // Write header
            writer.WriteUInt16BigEndian(SubtableFormat); // 1

            long markCoverageOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // MarkCoverage offset placeholder

            long baseCoverageOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // BaseCoverage offset placeholder

            writer.WriteUInt16BigEndian(MarkClassCount);

            long markArrayOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // MarkArray offset placeholder

            long baseArrayOffsetPos = writer.BaseStream.Position;
            writer.WriteUInt16BigEndian(0); // BaseArray offset placeholder

            // Write MarkArray
            long markArrayStart = writer.BaseStream.Position;
            ushort markArrayOffset = (ushort)(markArrayStart - subtableStart);

            if (MarkArray != null)
            {
                writer.WriteUInt16BigEndian(MarkArray.MarkCount);

                // Reserve space for anchor offsets
                long markAnchorOffsetsPos = writer.BaseStream.Position;
                for (int i = 0; i < MarkArray.MarkCount; i++)
                {
                    writer.WriteUInt16BigEndian(0); // MarkClass placeholder
                    writer.WriteUInt16BigEndian(0); // MarkAnchor offset placeholder
                }

                // Write MarkRecords and Anchors
                for (int i = 0; i < MarkArray.MarkCount; i++)
                {
                    var record = MarkArray.Records[i];

                    // Update MarkClass
                    long currentPos = writer.BaseStream.Position;
                    writer.BaseStream.Seek(markAnchorOffsetsPos + (i * 4), System.IO.SeekOrigin.Begin);
                    writer.WriteUInt16BigEndian(record.MarkClass);
                    writer.BaseStream.Seek(currentPos, System.IO.SeekOrigin.Begin);

                    // Write Anchor
                    if (record.MarkAnchor != null)
                    {
                        long anchorStart = writer.BaseStream.Position;
                        ushort anchorOffset = (ushort)(anchorStart - markArrayStart);

                        AnchorTableSerializer.Serialize(writer, record.MarkAnchor);

                        // Update anchor offset
                        currentPos = writer.BaseStream.Position;
                        writer.BaseStream.Seek(markAnchorOffsetsPos + (i * 4) + 2, System.IO.SeekOrigin.Begin);
                        writer.WriteUInt16BigEndian(anchorOffset);
                        writer.BaseStream.Seek(currentPos, System.IO.SeekOrigin.Begin);
                    }
                }
            }

            // Update MarkArray offset
            long resumePos = writer.BaseStream.Position;
            writer.BaseStream.Seek(markArrayOffsetPos, System.IO.SeekOrigin.Begin);
            writer.WriteUInt16BigEndian(markArrayOffset);
            writer.BaseStream.Seek(resumePos, System.IO.SeekOrigin.Begin);

            // Write BaseArray
            long baseArrayStart = writer.BaseStream.Position;
            ushort baseArrayOffset = (ushort)(baseArrayStart - subtableStart);

            if (BaseArray != null)
            {
                writer.WriteUInt16BigEndian(BaseArray.BaseCount);

                // Reserve space for anchor offsets
                long baseAnchorOffsetsPos = writer.BaseStream.Position;
                for (int i = 0; i < BaseArray.BaseCount; i++)
                {
                    for (int j = 0; j < MarkClassCount; j++)
                    {
                        writer.WriteUInt16BigEndian(0); // Placeholder
                    }
                }

                // Write BaseAnchors
                for (int i = 0; i < BaseArray.BaseCount; i++)
                {
                    var record = BaseArray.Records[i];

                    for (int j = 0; j < MarkClassCount; j++)
                    {
                        if (record.BaseAnchors != null && j < record.BaseAnchors.Length && record.BaseAnchors[j] != null)
                        {
                            long anchorStart = writer.BaseStream.Position;
                            ushort anchorOffset = (ushort)(anchorStart - baseArrayStart);

                            AnchorTableSerializer.Serialize(writer, record.BaseAnchors[j]);

                            // Update anchor offset
                            long currentPos = writer.BaseStream.Position;
                            writer.BaseStream.Seek(baseAnchorOffsetsPos + (i * MarkClassCount * 2) + (j * 2), System.IO.SeekOrigin.Begin);
                            writer.WriteUInt16BigEndian(anchorOffset);
                            writer.BaseStream.Seek(currentPos, System.IO.SeekOrigin.Begin);
                        }
                    }
                }
            }

            // Update BaseArray offset
            resumePos = writer.BaseStream.Position;
            writer.BaseStream.Seek(baseArrayOffsetPos, System.IO.SeekOrigin.Begin);
            writer.WriteUInt16BigEndian(baseArrayOffset);
            writer.BaseStream.Seek(resumePos, System.IO.SeekOrigin.Begin);

            // Write MarkCoverage table
            if (MarkCoverage != null)
            {
                ushort coverageOffset = (ushort)(writer.BaseStream.Position - subtableStart);
                resumePos = writer.BaseStream.Position;

                writer.BaseStream.Seek(markCoverageOffsetPos, System.IO.SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(coverageOffset);
                writer.BaseStream.Seek(resumePos, System.IO.SeekOrigin.Begin);

                MarkCoverage.Serialize(writer);
            }

            // Write BaseCoverage table
            if (BaseCoverage != null)
            {
                ushort coverageOffset = (ushort)(writer.BaseStream.Position - subtableStart);
                resumePos = writer.BaseStream.Position;

                writer.BaseStream.Seek(baseCoverageOffsetPos, System.IO.SeekOrigin.Begin);
                writer.WriteUInt16BigEndian(coverageOffset);
                writer.BaseStream.Seek(resumePos, System.IO.SeekOrigin.Begin);

                BaseCoverage.Serialize(writer);
            }
        }
    }
}