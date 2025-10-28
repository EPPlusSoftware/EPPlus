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
namespace EPPlus.Fonts.OpenType.Tables.Hmtx
{
    public class LongHorMetric
    {
        /// <summary>
        /// Advance width, in font design units.
        /// </summary>
        public ushort advanceWidth{ get; set; }

        /// <summary>
        /// Glyph left side bearing, in font design units.
        /// </summary>
        public short lsb { get; set; }

        public override string ToString()
        {
            return $"advanceWidth: {advanceWidth}, lsb: {lsb}";
        }
    }
}
