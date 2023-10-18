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
    /// Type of Trendline for a chart
    /// </summary>
    public enum eTrendLine
    {
        /// <summary>
        /// The trendline will be an exponential curve. y = abx
        /// </summary>
        Exponential,
        /// <summary>
        /// The trendline will be a linear curve. y = mx + b
        /// </summary>
        Linear,
        /// <summary>
        /// The trendline will be a logarithmic curve y = a log x + b
        /// </summary>
        Logarithmic,
        /// <summary>
        /// The trendline will be the moving average.
        /// </summary>
        MovingAverage,
        /// <summary>
        /// The trendline will be a polynomial curve.
        /// </summary>
        Polynomial,
        /// <summary>
        /// The trendline will be a power curve. y = axb
        /// </summary>
        Power
    }
}