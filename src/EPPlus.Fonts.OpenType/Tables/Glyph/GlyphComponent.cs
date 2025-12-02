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
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Glyph
{
    public class GlyphComponent : FontTableElement
    {
        public CompositeGlyphFlags Flags { get; set; }
        public ushort GlyphIndex { get; set; }
        public short Argument1 { get; set; }
        public short Argument2 { get; set; }

        // Transformation fields (F2Dot14 format)
        public short Scale { get; set; }
        public short XScale { get; set; }
        public short YScale { get; set; }
        public short Scale01 { get; set; }
        public short Scale10 { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            writer.WriteUInt16BigEndian((ushort)Flags);
            writer.WriteUInt16BigEndian(GlyphIndex);

            if ((Flags & CompositeGlyphFlags.ARG_1_AND_2_ARE_WORDS) != 0)
            {
                writer.WriteInt16BigEndian(Argument1);
                writer.WriteInt16BigEndian(Argument2);
            }
            else
            {
                writer.Write((byte)Argument1);
                writer.Write((byte)Argument2);
            }

            if ((Flags & CompositeGlyphFlags.WE_HAVE_A_SCALE) != 0)
            {
                writer.WriteInt16BigEndian(Scale);
            }
            else if ((Flags & CompositeGlyphFlags.WE_HAVE_AN_X_AND_Y_SCALE) != 0)
            {
                writer.WriteInt16BigEndian(XScale);
                writer.WriteInt16BigEndian(YScale);
            }
            else if ((Flags & CompositeGlyphFlags.WE_HAVE_A_TWO_BY_TWO) != 0)
            {
                writer.WriteInt16BigEndian(XScale);
                writer.WriteInt16BigEndian(Scale01);
                writer.WriteInt16BigEndian(Scale10);
                writer.WriteInt16BigEndian(YScale);
            }
        }
    }
}
