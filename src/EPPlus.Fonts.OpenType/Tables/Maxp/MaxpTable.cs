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
namespace EPPlus.Fonts.OpenType.Tables.Maxp
{
    public class MaxpTable : FontTableBase
    {
        /// <summary>
        /// 0x00005000 for version 0.5
        /// 0x00010000 for version 1.0.
        /// </summary>
        public Version16Dot16 version { get; set; }
        
        /// <summary>
        /// The number of glyphs in the font.
        /// </summary>
        public ushort numGlyphs { get; set; }

        /// <summary>
        /// Maximum points in a non-composite glyph.
        /// </summary>
        public ushort maxPoints { get; set; }

        /// <summary>
        /// Maximum contours in a non-composite glyph.
        /// </summary>
        public ushort maxContours { get; set; }

        /// <summary>
        /// Maximum points in a composite glyph.
        /// </summary>
        public ushort maxCompositePoints { get; set; }

        /// <summary>
        /// Maximum contours in a composite glyph.
        /// </summary>
        public ushort maxCompositeContours { get; set; }

        /// <summary>
        /// 1 if instructions do not use the twilight zone (Z0), or
        /// 2 if instructions do use Z0; should be set to 2 in
        /// most cases.
        /// </summary>
        public ushort maxZones { get; set; }

        /// <summary>
        /// Maximum points used in Z0.
        /// </summary>
        public ushort maxTwilightPoints { get; set; }

        /// <summary>
        /// Number of Storage Area locations.
        /// </summary>
        public ushort maxStorage { get; set; }

        /// <summary>
        /// Number of FDEFs, equal to the highest function number + 1.
        /// </summary>
        public ushort maxFunctionDefs { get; set; }

        /// <summary>
        /// Number of IDEFs.
        /// </summary>
        public ushort maxInstructionDefs { get; set; }

        public ushort maxStackElements { get; set; }

        public ushort maxSizeOfInstructions { get; set; }

        public ushort maxComponentElements { get; set; }

        public ushort maxComponentDepth { get; set; }


        internal override void SerializeInternal(FontsBinaryWriter writer)
        {
            version.Serialize(writer);
            writer.WriteUInt16BigEndian(numGlyphs);
            if(version.Major == 1)
            {
                writer.WriteUInt16BigEndian(maxPoints);
                writer.WriteUInt16BigEndian(maxContours);
                writer.WriteUInt16BigEndian(maxCompositePoints);
                writer.WriteUInt16BigEndian(maxCompositeContours);
                writer.WriteUInt16BigEndian(maxZones);
                writer.WriteUInt16BigEndian(maxTwilightPoints);
                writer.WriteUInt16BigEndian(maxStorage);
                writer.WriteUInt16BigEndian(maxFunctionDefs);
                writer.WriteUInt16BigEndian(maxInstructionDefs);
                writer.WriteUInt16BigEndian(maxStackElements);
                writer.WriteUInt16BigEndian(maxSizeOfInstructions);
                writer.WriteUInt16BigEndian(maxComponentElements);
                writer.WriteUInt16BigEndian(maxComponentDepth);
            }
        }
    }
}
