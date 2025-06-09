/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2025         EPPlus Software AB           Initial release EPPlus 8
 *************************************************************************************************/
namespace OfficeOpenXml.Drawing.Shape
{
    internal class ShapeGuidePoint
    {
        public static implicit operator ShapeGuidePoint(int value)
        {
            return new ShapeGuidePoint(value);
        }
        public static implicit operator int(ShapeGuidePoint value)
        {
            return value.Value;
        }

        internal static string valuePrefix = "val ";
        private int _value;
        internal string fmlaValue
        {
            get { return valuePrefix + _value; }
            private set { }
        }
        internal int Value
        {
            get { return _value; }
            set { _value = value; }
        }

        internal ShapeGuidePoint(int value)
        {
            this._value = value;
        }
    }
}
