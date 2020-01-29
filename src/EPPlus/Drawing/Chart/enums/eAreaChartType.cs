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
    /// Area chart type
    /// </summary>
    public enum eAreaChartType
    {
        /// <summary>
        /// An area chart
        /// </summary>
        Area = 1,
        /// <summary>
        /// A stacked area chart
        /// </summary>
        AreaStacked = 76,
        /// <summary>
        /// A stacked 100 percent area chart
        /// </summary>
        AreaStacked100 = 77,
        /// <summary>
        /// An 3D area chart
        /// </summary>
        Area3D = -4098,
        /// <summary>
        /// A stacked area 3D chart
        /// </summary>
        AreaStacked3D = 78,
        /// <summary>
        /// A stacked 100 percent 3D area chart
        /// </summary>
        AreaStacked1003D = 79,
    }
}