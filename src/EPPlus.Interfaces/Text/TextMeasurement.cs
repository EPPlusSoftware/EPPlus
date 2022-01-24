/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  1/4/2021         EPPlus Software AB           EPPlus Interfaces 1.0
 *************************************************************************************************/

namespace OfficeOpenXml.Interfaces.Text
{
    public struct TextMeasurement
    {
        public TextMeasurement(float width, float height)
        {
            Width = width;
            Height = height;
        }

        /// <summary>
        /// Width of the text
        /// </summary>
        public float Width { get; set; }

        /// <summary>
        /// Height of the text
        /// </summary>
        public float Height { get; set; }

        public static TextMeasurement Empty
        {
            get { return new TextMeasurement(-1, -1); }
        }

        /// <summary>
        /// Returns true if this is an empty measurement
        /// </summary>
        public bool IsEmpty
        {
            get { return Width == -1 && Height == -1; }
        }
    }
}
