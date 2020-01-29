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
namespace OfficeOpenXml.Table
{
    /// <summary>
    /// Build-in table row functions
    /// </summary>
    public enum RowFunctions
    {
        /// <summary>
        /// Average
        /// </summary>
        Average,
        /// <summary>
        /// Count
        /// </summary>
        Count,
        /// <summary>
        /// Count, numbers
        /// </summary>
        CountNums,
        /// <summary>
        /// Custum function
        /// </summary>
        Custom,
        /// <summary>
        /// Maximum
        /// </summary>
        Max,
        /// <summary>
        /// Minimum
        /// </summary>
        Min,
        /// <summary>
        /// None
        /// </summary>
        None,
        /// <summary>
        /// Standard deviation
        /// </summary>
        StdDev,
        /// <summary>
        /// Summary
        /// </summary>
        Sum,
        /// <summary>
        /// Variation
        /// </summary>
        Var
    }
}