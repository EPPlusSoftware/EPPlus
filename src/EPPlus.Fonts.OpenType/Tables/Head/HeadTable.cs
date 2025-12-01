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
namespace EPPlus.Fonts.OpenType.Tables.Head
{
    /// <summary>
    /// This table gives global information about the font.
    /// </summary>
    public class HeadTable : FontTableBase
    {
        public enum IndexToLocFormats : short
        {
            Offset16 = 0,
            Offset32 = 1
        }

        public override string Name => TableNames.Head;

        override public bool IsEssentialTable => true;

        public ushort MajorVersion { get; set; }

        public ushort MinorVersion { get; set; }

        /// <summary>
        /// Set by font manufacturer.
        /// </summary>
        public int FontRevision { get; set; }

        /// <summary>
        /// To compute: set it to 0, sum the entire font as uint32,then store 0xB1B0AFBA - sum. 
        /// If the font is used as acomponent in a font collection file, the value of thisfield 
        /// will be invalidated by changes to the file structure and font table directory, and must be ignored.
        /// </summary>
        public uint ChecksumAdjustment { get; set; }

        /// <summary>
        /// Set to 0x5F0F3CF5.
        /// </summary>
        public uint MagicNumber { get; set; }

        public ushort Flags { get; set; }

        /// <summary>
        /// Set to a value from 16 to 16384. Any value in this range is valid. In fonts that have TrueType outlines, a power of 2 is recommended as this allows performance optimizations in some rasterizers.
        /// </summary>
        public ushort UnitsPerEm { get; set; }

        /// <summary>
        /// Number of seconds since 12:00 midnight that startedJanuary 1st, 1904, in GMT/UTC time zone.
        /// </summary>
        public long Created { get; set; }

        /// <summary>
        /// Number of seconds since 12:00 midnight that startedJanuary 1st, 1904, in GMT/UTC time zone.
        /// </summary>
        public long Modified { get; set; }

        /// <summary>
        /// Minimum x coordinate across all glyph bounding boxes.
        /// </summary>
        public short Xmin { get; set; }

        /// <summary>
        /// Minimum y coordinate across all glyph bounding boxes.
        /// </summary>
        public short Ymin { get; set; }

        /// <summary>
        /// Maximum x coordinate across all glyph bounding boxes.
        /// </summary>
        public short Xmax { get; set; }

        /// <summary>
        /// Maximum y coordinate across all glyph bounding boxes.
        /// </summary>
        public short Ymax { get; set; }

        /// <summary>
        /// Bit 0: Bold(if set to 1);
        /// Bit 1: Italic(if set to 1)
        /// Bit 2: Underline(if set to 1)
        /// Bit 3: Outline(if set to 1)
        /// Bit 4: Shadow(if set to 1)
        /// Bit 5: Condensed(if set to 1)
        /// Bit 6: Extended(if set to 1)
        /// Bits 7 – 15: Reserved(set to 0)
        /// </summary>
        public ushort MacStyle { get; set; }

        /// <summary>
        /// Smallest readable size in pixels.
        /// </summary>
        public ushort LowestRecPPEM { get; set; }

        /// <summary>
        /// Deprecated (Set to 2).
        /// 0: Fully mixed directional glyphs;
        /// 1: Only strongly left to right;
        /// 2: Like 1 but also contains neutrals;
        /// -1: Only strongly right to left;
        /// -2: Like -1 but also contains neutrals.
        /// </summary>
        public short FontDirectionHint { get; set; }


        /// <summary>
        /// 0 for short offsets (Offset16), 1 for long (Offset32).
        /// </summary>
        public IndexToLocFormats IndexToLocFormat { get; set; }

        /// <summary>
        /// 0 for current format.
        /// </summary>
        public short GlyphDataFormat { get; set; }

        public BoundingRectangle GetDefaultBounds()
        {
            return new BoundingRectangle(Xmin, Ymin, Xmax, Ymax);
        }

        internal override void SerializeInternal(FontsBinaryWriter writer)
        {
            writer.WriteUInt16BigEndian(MajorVersion);
            writer.WriteUInt16BigEndian(MinorVersion);
            writer.WriteInt32BigEndian(FontRevision); // Fixed32: 16.16 format, men du har int – ev. konvertering behövs
            writer.WriteUInt32BigEndian(ChecksumAdjustment);
            writer.WriteUInt32BigEndian(MagicNumber);
            writer.WriteUInt16BigEndian(Flags);
            writer.WriteUInt16BigEndian(UnitsPerEm);
            writer.WriteInt64BigEndian(Created);
            writer.WriteInt64BigEndian(Modified);
            writer.WriteInt16BigEndian(Xmin);
            writer.WriteInt16BigEndian(Ymin);
            writer.WriteInt16BigEndian(Xmax);
            writer.WriteInt16BigEndian(Ymax);
            writer.WriteUInt16BigEndian(MacStyle);
            writer.WriteUInt16BigEndian(LowestRecPPEM);
            writer.WriteInt16BigEndian(FontDirectionHint);
            writer.WriteInt16BigEndian((short)IndexToLocFormat);
            writer.WriteInt16BigEndian(GlyphDataFormat);
        }

        internal override void Clear()
        {
            throw new System.NotImplementedException();
        }

        public HeadTable Clone()
        {
            return new HeadTable
            {
                MajorVersion = this.MajorVersion,
                MinorVersion = this.MinorVersion,
                FontRevision = this.FontRevision,
                ChecksumAdjustment = this.ChecksumAdjustment,
                MagicNumber = this.MagicNumber,
                Flags = this.Flags,
                UnitsPerEm = this.UnitsPerEm,
                Created = this.Created,
                Modified = this.Modified,
                Xmin = this.Xmin,
                Ymin = this.Ymin,
                Xmax = this.Xmax,
                Ymax = this.Ymax,
                MacStyle = this.MacStyle,
                LowestRecPPEM = this.LowestRecPPEM,
                FontDirectionHint = this.FontDirectionHint,
                IndexToLocFormat = this.IndexToLocFormat,
                GlyphDataFormat = this.GlyphDataFormat
            };
        }
    }

    
}
