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
namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Build-in table row functions
    /// </summary>
    public enum DataFieldFunctions
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
        /// Max value
        /// </summary>
        Max,
        /// <summary>
        /// Minimum value
        /// </summary>
        Min,
        /// <summary>
        /// The product
        /// </summary>
        Product,
        /// <summary>
        /// None
        /// </summary>
        None,
        /// <summary>
        /// Standard deviation
        /// </summary>
        StdDev,
        /// <summary>
        /// Standard deviation of a population,
        /// </summary>
        StdDevP,
        /// <summary>
        /// Sum
        /// </summary>
        Sum,
        /// <summary>
        /// Variation
        /// </summary>
        Var,
        /// <summary>
        /// The variance of a population
        /// </summary>
        VarP
    }
}