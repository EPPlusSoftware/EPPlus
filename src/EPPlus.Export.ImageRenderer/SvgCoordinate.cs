/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/

namespace EPPlusImageRenderer
{
    public struct SvgCoordinate
    {
        public static implicit operator SvgCoordinate(short value)
        {
            return new SvgCoordinate(value);
        }
        public static implicit operator short(SvgCoordinate value)
        {
            return value.Value;
        }
        public SvgCoordinate(short value)
        {
            Origin = default;
            Value = value;
            PointName = default;
            Type = default;
        }
        internal short Origin { get; set; }
        internal short Value { get; set; }
        internal string PointName { get; set; }
        internal AdjustmentPointType Type { get; set; }
        public override int GetHashCode()
        {
            return Value.GetHashCode();
        }
        public override bool Equals(object obj)
        {
            if (obj is SvgCoordinate c)
            {
                return Value.Equals(c.Value);
            }
            return base.Equals(obj);
        }

    }
}