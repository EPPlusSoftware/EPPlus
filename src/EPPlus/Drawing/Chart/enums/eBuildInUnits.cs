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
    /// Build in units for a chart axis
    /// </summary>
    public enum eBuildInUnits : long
    {
        /// <summary>
        /// 100
        /// </summary>
        hundreds = 100,
        /// <summary>
        /// 1,000
        /// </summary>
        thousands = 1000,
        /// <summary>
        /// 10,000
        /// </summary>
        tenThousands = 10000,
        /// <summary>
        /// 100,000
        /// </summary>
        hundredThousands = 100000,
        /// <summary>
        /// 1,000,000
        /// </summary>
        millions = 1000000,
        /// <summary>
        /// 10,000,000
        /// </summary>
        tenMillions = 10000000,
        /// <summary>
        /// 10,000,000
        /// </summary>
        hundredMillions = 100000000,
        /// <summary>
        /// 1,000,000,000
        /// </summary>
        billions = 1000000000,
        /// <summary>
        /// 1,000,000,000,000
        /// </summary>
        trillions = 1000000000000
    }
}