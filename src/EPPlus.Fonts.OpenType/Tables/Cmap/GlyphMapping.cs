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
using System.Diagnostics;


namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    [DebuggerDisplay("{CharacterCode} - '{Char}': {GlyphIndex}")]
    public class GlyphMapping : FontTableElement
    {
        public ushort CharacterCode { get; set; }

        public ushort GlyphIndex { get; set; }

        public char Char => Convert.ToChar(CharacterCode);

        public override string ToString()
        {
            return Char.ToString() + ": " + GlyphIndex;
        }


        internal override void Serialize(FontsBinaryWriter writer)
        {
            writer.WriteUInt16BigEndian(CharacterCode);
            writer.WriteUInt16BigEndian(GlyphIndex);
        }

    }
}
