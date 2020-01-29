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
    /// Line chart type
    /// </summary>
    public enum eLineChartType
    {
        /// <summary>
        /// A 3D line chart
        /// </summary>
        Line3D = -4101,
        /// <summary>
        /// A line chart
        /// </summary>
        Line = 4,
        /// <summary>
        /// A line chart with markers
        /// </summary>
        LineMarkers = 65,
        /// <summary>
        /// A stacked line chart with markers
        /// </summary>
        LineMarkersStacked = 66,
        /// <summary>
        /// A 100% stacked line chart with markers
        /// </summary>
        LineMarkersStacked100 = 67,
        /// <summary>
        /// A stacked line chart
        /// </summary>
        LineStacked = 63,
        /// <summary>
        /// A 100% stacked line chart
        /// </summary>
        LineStacked100 = 64,
    }
}