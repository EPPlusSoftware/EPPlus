/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Where the axis cross. 
    /// </summary>
    public enum eCrosses
    {
        /// <summary>
        /// The category axis crosses at the zero point of the valueaxis or the lowest or higest value if scale is over or below zero.
        /// </summary>
        AutoZero,
        /// <summary>
        /// The axis crosses at the maximum value
        /// </summary>
        Max,
        /// <summary>
        /// Axis crosses at the minimum value
        /// </summary>
        Min
    }
}