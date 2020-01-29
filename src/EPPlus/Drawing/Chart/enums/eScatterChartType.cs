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
    /// Scatter chart types
    /// </summary>
    public enum eScatterChartType
    {
        /// <summary>
        /// A XY scatter chart
        /// </summary>
        XYScatter = -4169,
        /// <summary>
        /// A scatter line chart with markers
        /// </summary>
        XYScatterLines = 74,
        /// <summary>
        /// A scatter line chart with no markers
        /// </summary>
        XYScatterLinesNoMarkers = 75,
        /// <summary>
        /// A scatter line chart with markers and smooth lines
        /// </summary>
        XYScatterSmooth = 72,
        /// <summary>
        /// A scatter line chart with no markers and smooth lines
        /// </summary>
        XYScatterSmoothNoMarkers = 73
    }
}