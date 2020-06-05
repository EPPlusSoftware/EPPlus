/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    04/16/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// Side positions for a chart element
    /// </summary>
    public enum eNumericDataType
    {
        /// <summary>
        /// The dimension is a value.
        /// </summary>
        Value,
        /// <summary>
        /// The dimension is an x-coordinate.
        /// </summary>
        X,
        /// <summary>
        /// The dimension is a y-coordinate.
        /// </summary>
        Y,
        /// <summary>
        /// The dimension is a size.
        /// </summary>
        Size,
        /// <summary>
        /// The dimension is a value determining a color.        
        /// </summary>
        ColorValue
    }
}