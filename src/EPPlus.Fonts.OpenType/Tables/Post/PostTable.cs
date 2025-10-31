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
namespace EPPlus.Fonts.OpenType.Tables.Post
{
    public class PostTable
    {
        public int version { get; set; }
        public double italicAngle { get; set; }
        public short underlinePosition {  get; set; }
        public short underlineThickness { get; set; }
        public uint isFixedPitch { get; set; }
        public uint minMemType42 { get; set; }
        public uint maxMemType42 { get; set; }
        public uint minMemType1 { get; set; }
        public uint maxMemType1 { get;set; }
    }
}
