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
    public class CompositeGlyph : FontTableElement
    {
        public List<GlyphComponent> Components { get; set; } = new List<GlyphComponent>();
        public byte[] Instructions { get; set; } = new byte[0];

        internal override void Serialize(FontsBinaryWriter writer)
        {
            foreach (var component in Components)
            {
                component.Serialize(writer);
            }

            // Kontrollera om sista komponenten har WE_HAVE_INSTRUCTIONS-flaggan
            if ((Components.Last().Flags & CompositeGlyphFlags.WE_HAVE_INSTRUCTIONS) != 0)
            {
                writer.WriteUInt16BigEndian((ushort)Instructions.Length);
                writer.Write(Instructions);
            }
        }
    }
}
