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
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Glyph.Serialization
{
    internal static class CompositeGlyphDeserializer
    {
        public static CompositeGlyph Deserialize(FontsBinaryReader reader)
        {
            var glyph = new CompositeGlyph();
            bool moreComponents;

            do
            {
                var flags = (CompositeGlyphFlags)reader.ReadUInt16BigEndian();
                var glyphIndex = reader.ReadUInt16BigEndian();

                short arg1, arg2;
                if ((flags & CompositeGlyphFlags.ARG_1_AND_2_ARE_WORDS) != 0)
                {
                    arg1 = reader.ReadInt16BigEndian();
                    arg2 = reader.ReadInt16BigEndian();
                }
                else
                {
                    arg1 = reader.ReadByte();
                    arg2 = reader.ReadByte();
                }

                var component = new GlyphComponent
                {
                    Flags = flags,
                    GlyphIndex = glyphIndex,
                    Argument1 = arg1,
                    Argument2 = arg2
                };

                if ((flags & CompositeGlyphFlags.WE_HAVE_A_SCALE) != 0)
                {
                    component.Scale = reader.ReadInt16BigEndian();
                }
                else if ((flags & CompositeGlyphFlags.WE_HAVE_AN_X_AND_Y_SCALE) != 0)
                {
                    component.XScale = reader.ReadInt16BigEndian();
                    component.YScale = reader.ReadInt16BigEndian();
                }
                else if ((flags & CompositeGlyphFlags.WE_HAVE_A_TWO_BY_TWO) != 0)
                {
                    component.XScale = reader.ReadInt16BigEndian();
                    component.Scale01 = reader.ReadInt16BigEndian();
                    component.Scale10 = reader.ReadInt16BigEndian();
                    component.YScale = reader.ReadInt16BigEndian();
                }

                glyph.Components.Add(component);
                moreComponents = (flags & CompositeGlyphFlags.MORE_COMPONENTS) != 0;

            } while (moreComponents);

            if ((glyph.Components.Last().Flags & CompositeGlyphFlags.WE_HAVE_INSTRUCTIONS) != 0)
            {
                var instructionLength = reader.ReadUInt16BigEndian();
                glyph.Instructions = reader.ReadBytes(instructionLength);
            }

            return glyph;
        }
    }
}
