/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/07/2026         EPPlus Software AB           GPOS PairValueRecord
 *************************************************************************************************/

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2
{
    /// <summary>
    /// Single pair value record
    /// </summary>
    public class PairValueRecord
    {
        /// <summary>
        /// Glyph ID of second glyph in pair
        /// </summary>
        public ushort SecondGlyph { get; set; }

        /// <summary>
        /// Positioning for first glyph
        /// </summary>
        public ValueRecord Value1 { get; set; }

        /// <summary>
        /// Positioning for second glyph
        /// </summary>
        public ValueRecord Value2 { get; set; }

        internal void Serialize(FontsBinaryWriter writer, ushort valueFormat1, ushort valueFormat2)
        {
            writer.WriteUInt16BigEndian(SecondGlyph);
            Value1?.Write(writer, valueFormat1);
            Value2?.Write(writer, valueFormat2);
        }
    }
}
