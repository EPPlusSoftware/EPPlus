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
    /// Pie and Doughnut chart type
    /// </summary>
    public enum ePieChartType
    {
        /// <summary>
        /// A pie chart
        /// </summary>
        Pie = 5,
        /// <summary>
        /// An exploded pie chart
        /// </summary>
        PieExploded = 69,
        /// <summary>
        /// A 3D pie chart
        /// </summary>
        Pie3D = -4102,
        /// <summary>
        /// A exploded 3D pie chart
        /// </summary>
        PieExploded3D = 70
    }
}