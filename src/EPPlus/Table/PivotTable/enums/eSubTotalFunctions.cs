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
using System;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Built-in subtotal functions
    /// </summary>
    [Flags]
    public enum eSubTotalFunctions
    {
        /// <summary>
        /// None
        /// </summary>
        None = 1,
        /// <summary>
        /// Count cells that are numbers.
        /// </summary>
        Count = 2,
        /// <summary>
        /// Count cells that are not empty.
        /// </summary>
        CountA = 4,
        /// <summary>
        /// Average
        /// </summary>
        Avg = 8,
        /// <summary>
        /// Default, total
        /// </summary>
        Default = 16,
        /// <summary>
        /// Minimum
        /// </summary>
        Min = 32,
        /// <summary>
        /// Maximum
        /// </summary>
        Max = 64,
        /// <summary>
        /// Product
        /// </summary>
        Product = 128,
        /// <summary>
        /// Standard deviation
        /// </summary>
        StdDev = 256,
        /// <summary>
        /// Standard deviation of a population
        /// </summary>
        StdDevP = 512,
        /// <summary>
        /// Summary
        /// </summary>
        Sum = 1024,
        /// <summary>
        /// Variation
        /// </summary>
        Var = 2048,
        /// <summary>
        /// Variation of a population
        /// </summary>
        VarP = 4096
    }
}