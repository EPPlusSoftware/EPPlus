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

namespace EPPlus.Fonts.OpenType.Tables.Post
{
    public class PostTable : FontTableBase
    {
        public Version16Dot16 version { get; set; }
        public Fixed16Dot16 italicAngle { get; set; }
        public short underlinePosition {  get; set; }
        public short underlineThickness { get; set; }
        public uint isFixedPitch { get; set; }
        public uint minMemType42 { get; set; }
        public uint maxMemType42 { get; set; }
        public uint minMemType1 { get; set; }
        public uint maxMemType1 { get;set; }

        // version 2

        public ushort numGlyphs {  get; set; }
        public ushort[] glyphNameIndex { get; set; }

        public string[] glyphNames { get; set; }

        internal override void SerializeInternal(FontsBinaryWriter writer)
        {
            version.Serialize(writer);
            italicAngle.Serialize(writer);
            writer.WriteInt16BigEndian(underlinePosition);
            writer.WriteInt16BigEndian(underlineThickness);
            writer.WriteUInt32BigEndian(isFixedPitch);
            writer.WriteUInt32BigEndian(minMemType42);
            writer.WriteUInt32BigEndian(maxMemType42);
            writer.WriteUInt32BigEndian(minMemType1);
            writer.WriteUInt32BigEndian(maxMemType1);
            if(version.Major == 2 && version.Minor == 0)
            {
                writer.WriteUInt16BigEndian(numGlyphs);

                // 1. Write glyphNameIndex[]
                foreach (var index in glyphNameIndex)
                {
                    writer.WriteUInt16BigEndian(index);
                }

                // 2. Write Pascal strings for custom names (index >= 258)
                for (int i = 0; i < glyphNameIndex.Length; i++)
                {
                    ushort index = glyphNameIndex[i];
                    if (index >= 258)
                    {
                        string name = glyphNames[i];
                        byte[] nameBytes = System.Text.Encoding.ASCII.GetBytes(name);
                        if (nameBytes.Length > 255)
                            throw new InvalidOperationException($"Glyph name '{name}' is too long (max 255 bytes).");

                        writer.Write((byte)nameBytes.Length); // Pascal length
                        writer.Write(nameBytes);              // ASCII name
                    }
                }
            }
        }
    }
}
