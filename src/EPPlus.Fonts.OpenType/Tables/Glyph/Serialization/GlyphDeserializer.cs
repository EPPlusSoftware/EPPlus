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
namespace EPPlus.Fonts.OpenType.Tables.Glyph.Serialization
{
    internal static class GlyphDeserializer
    {
        public static Glyph Deserialize(FontsBinaryReader reader)
        {
            var numberOfContours = reader.ReadInt16BigEndian();
            var xMin = reader.ReadInt16BigEndian();
            var yMin = reader.ReadInt16BigEndian();
            var xMax = reader.ReadInt16BigEndian();
            var yMax = reader.ReadInt16BigEndian();

            var glyph = new Glyph
            {
                Header = new GlyphHeader
                {
                    numberOfContours = numberOfContours,
                    xMin = xMin,
                    yMin = yMin,
                    xMax = xMax,
                    yMax = yMax
                }
            };

            if (numberOfContours > 0)
            {
                glyph.SimpleData = SimpleGlyphDeserializer.Deserialize(reader, numberOfContours);
            }
            else if (numberOfContours < 0)
            {
                glyph.CompositeData = CompositeGlyphDeserializer.Deserialize(reader);
            }

            return glyph;
        }
    }
}
