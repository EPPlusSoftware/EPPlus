/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/22/2020         EPPlus Software AB       Added this class
 *************************************************************************************************/
namespace OfficeOpenXml
{
    /// <summary>
    /// How the axis label should be alignted within the major tickmarks.
    /// </summary>
    public enum eAxisLabelAlignment
    {
        /// <summary>
        /// The text shall be centered
        /// </summary>
        Left,
        /// <summary>
        /// The text shall be left justified.
        /// </summary>
        Center,
        /// <summary>
        /// The text shall be right justified.
        /// </summary>
        Right
    }
}