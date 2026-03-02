/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/18/2026         EPPlus Software AB           vhea table implementation (vertical text support)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables;

namespace EPPlus.Fonts.OpenType.Tables.Vhea
{
    /// <summary>
    /// Represents the OpenType 'vhea' (Vertical Header) table.
    /// Contains global metrics for vertically laid-out fonts (primarily CJK).
    /// Structurally identical to 'hhea' but for the vertical axis.
    /// Source: https://learn.microsoft.com/en-us/typography/opentype/spec/vhea
    /// </summary>
    public class VheaTable : FontTableBase
    {
        public override string Name => TableNames.Vhea;

        /// <summary>
        /// Optional table - only present in fonts with vertical metrics (primarily CJK fonts).
        /// </summary>
        public override bool IsEssentialTable => false;

        // --- Version ---

        /// <summary>
        /// Table version. Either 0x00011000 (version 1.1) or 0x00010000 (version 1.0).
        /// Version 1.1 renamed some fields; the binary layout is identical.
        /// </summary>
        public uint Version { get; set; }

        // --- Vertical metrics (version 1.1 names; version 1.0 names in comments) ---

        /// <summary>
        /// The vertical typographic ascender for the font. (v1.0: ascent)
        /// Typically a positive value equal to half the em square.
        /// </summary>
        public short Ascent { get; set; }

        /// <summary>
        /// The vertical typographic descender for the font. (v1.0: descent)
        /// Typically a negative value equal to minus half the em square.
        /// </summary>
        public short Descent { get; set; }

        /// <summary>
        /// Additional line spacing for vertical layout. (v1.0: lineGap)
        /// </summary>
        public short LineGap { get; set; }

        /// <summary>
        /// Maximum advance height across all glyphs in the font.
        /// </summary>
        public short AdvanceHeightMax { get; set; }

        /// <summary>
        /// Minimum top side bearing across all glyphs with contours.
        /// </summary>
        public short MinTopSideBearing { get; set; }

        /// <summary>
        /// Minimum bottom side bearing across all glyphs with contours.
        /// Calculated as: advanceHeight - (tsb + yMin - yMax)
        /// </summary>
        public short MinBottomSideBearing { get; set; }

        /// <summary>
        /// Maximum extent: max(tsb + (yMax - yMin)) across all glyphs with contours.
        /// </summary>
        public short YMaxExtent { get; set; }

        /// <summary>
        /// The rise of the caret slope for vertical text (used to draw the caret).
        /// For vertical text: typically 0.
        /// </summary>
        public short CaretSlopeRise { get; set; }

        /// <summary>
        /// The run of the caret slope for vertical text.
        /// For vertical text: typically 1.
        /// </summary>
        public short CaretSlopeRun { get; set; }

        /// <summary>
        /// Offset of the caret from the glyph origin. Set to 0 for non-slanted fonts.
        /// </summary>
        public short CaretOffset { get; set; }

        // Reserved fields (must be 0 per spec)
        public short Reserved1 { get; set; }
        public short Reserved2 { get; set; }
        public short Reserved3 { get; set; }
        public short Reserved4 { get; set; }

        /// <summary>
        /// Set to 0. (Kept for future use per spec.)
        /// </summary>
        public short MetricDataFormat { get; set; }

        /// <summary>
        /// Number of LongVerMetric entries in the 'vmtx' table.
        /// Controls how vmtx is parsed - analogous to hhea.numberOfHMetrics.
        /// </summary>
        public ushort NumberOfVMetrics { get; set; }

        // --- Clone ---

        /// <summary>
        /// Creates a deep copy of this VheaTable.
        /// Used during font subsetting - NumberOfVMetrics will be updated by VheaSubsetProcessor.
        /// </summary>
        public VheaTable Clone()
        {
            return new VheaTable
            {
                Version = Version,
                Ascent = Ascent,
                Descent = Descent,
                LineGap = LineGap,
                AdvanceHeightMax = AdvanceHeightMax,
                MinTopSideBearing = MinTopSideBearing,
                MinBottomSideBearing = MinBottomSideBearing,
                YMaxExtent = YMaxExtent,
                CaretSlopeRise = CaretSlopeRise,
                CaretSlopeRun = CaretSlopeRun,
                CaretOffset = CaretOffset,
                Reserved1 = Reserved1,
                Reserved2 = Reserved2,
                Reserved3 = Reserved3,
                Reserved4 = Reserved4,
                MetricDataFormat = MetricDataFormat,
                NumberOfVMetrics = NumberOfVMetrics
            };
        }

        // --- Serialization ---

        internal override void SerializeInternal(FontsBinaryWriter writer, FontSerializationContext context)
        {
            writer.WriteUInt32BigEndian(Version);
            writer.WriteInt16BigEndian(Ascent);
            writer.WriteInt16BigEndian(Descent);
            writer.WriteInt16BigEndian(LineGap);
            writer.WriteInt16BigEndian(AdvanceHeightMax);
            writer.WriteInt16BigEndian(MinTopSideBearing);
            writer.WriteInt16BigEndian(MinBottomSideBearing);
            writer.WriteInt16BigEndian(YMaxExtent);
            writer.WriteInt16BigEndian(CaretSlopeRise);
            writer.WriteInt16BigEndian(CaretSlopeRun);
            writer.WriteInt16BigEndian(CaretOffset);
            writer.WriteInt16BigEndian(Reserved1);
            writer.WriteInt16BigEndian(Reserved2);
            writer.WriteInt16BigEndian(Reserved3);
            writer.WriteInt16BigEndian(Reserved4);
            writer.WriteInt16BigEndian(MetricDataFormat);
            writer.WriteUInt16BigEndian(NumberOfVMetrics);
        }

        internal override void Clear()
        {
            // Not used in current architecture
        }
    }
}