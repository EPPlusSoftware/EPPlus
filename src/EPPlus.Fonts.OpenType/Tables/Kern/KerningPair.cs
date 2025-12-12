using System;
using System.Diagnostics;

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
namespace EPPlus.Fonts.OpenType.Tables.Kern
{
    [DebuggerDisplay("l: {left}, r: {right}, v: {value}")]
    public class KerningPair
    {
        internal KerningPair(FontsBinaryReader reader)
        {
            left = reader.ReadUInt16BigEndian();
            right = reader.ReadUInt16BigEndian();
            value = reader.ReadInt16BigEndian();

            //Left is high-order/most significant
            //but since big-endian we must do it like this.
            Combined = ((uint)left << 16) | right;
        }

        /// <summary>
        /// Glyph-index for left glyph
        /// </summary>
        public ushort left { get; set; }

        /// <summary>
        /// Glyph-index for right glyph
        /// </summary>
        public ushort right { get; set; }
        /// <summary>
        /// FWORD in font-design units. 
        /// Negative value moves chars closer
        /// Positive value moves chars further apart
        /// </summary>
        public short value { get; set; }

        /// <summary>
        /// The pairs are ordered in numeric order based on the combinded uint32 value of left and right.
        /// (left is high order word) 
        /// Therefore useful in binary search later to have this number readily available
        /// Source: https://learn.microsoft.com/en-us/typography/opentype/spec/kern#format-0
        /// </summary>
        public uint Combined { get; set; }
    }
}
