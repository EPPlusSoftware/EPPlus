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
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// The ways to determine the length of the error bars
    /// </summary>
    public enum eErrorValueType
    {
        /// <summary>
        /// The length of the error bars will be determined by the Plus and Minus properties.
        /// </summary>
        Custom,
        /// <summary>
        /// The length of the error bars will be the fixed value determined by Error Bar Value property.
        /// </summary>
        FixedValue,
        /// <summary>
        /// The length of the error bars will be Error Bar Value percent of the data.
        /// </summary>
        Percentage,
        /// <summary>
        /// The length of the error bars will be Error Bar Value standard deviations of the data.
        /// </summary>
        StandardDeviation,
        /// <summary>
        /// The length of the error bars will be Error Bar Value standard errors of the data.
        /// </summary>
        StandardError
    }
}
