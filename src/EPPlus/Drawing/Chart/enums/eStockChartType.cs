/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/25/2020         EPPlus Software AB       Added this enum
 *************************************************************************************************/
 namespace OfficeOpenXml.Drawing.Chart
{
    public enum eStockChartType
    {
        /// <summary>
        /// Stock chart with a High, Low and Close serie
        /// </summary>
        StockHLC = 88,
        /// <summary>
        /// Stock chart with an Open, High, Low and Close serie
        /// </summary>
        StockOHLC = 89,
        /// <summary>
        /// Stock chart with an Volume, High, Low and Close serie
        /// </summary>
        StockVHLC = 90,
        /// <summary>
        /// Stock chart with an Volume, Open, High, Low and Close serie
        /// </summary>
        StockVOHLC = 91,
    }
}
