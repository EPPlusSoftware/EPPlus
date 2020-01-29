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
    /// How the series are grouped
    /// </summary>
    public enum eGrouping
    {
        /// <summary>
        /// Standard grouping
        /// </summary>
        Standard,
        /// <summary>
        /// Clustered grouping
        /// </summary>
        Clustered,
        /// <summary>
        /// Stacked grouping
        /// </summary>
        Stacked,
        /// <summary>
        /// 100% stacked grouping
        /// </summary>
        PercentStacked
    }
}