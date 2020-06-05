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
    /// The color type for a region map charts color variation
    /// </summary>
    public enum eColorValuePositionType
    {
        /// <summary>
        /// The position’s location on the gradient is determined the numerical value in the <see cref="ExcelChartExValueColor.PositionValue"/> property.
        /// </summary>
        Number,
        /// <summary>
        /// The position’s location on the gradient is determined by a fixed percent value in the <see cref="ExcelChartExValueColor.PositionValue"/> property, represented by the gradient. Ranges from 1 to 100 percent.
        /// </summary>
        Percent,
        /// <summary>
        /// The position is the minimum or maximum stop of the gradient.
        /// </summary>
        Extreme
    }
}