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
using System.Linq;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Represent strategies for handling precision and rounding of float/double values when calculating formulas.
    /// </summary>
    public enum PrecisionAndRoundingStrategy
    {
        /// <summary>
        /// Use .NET's default functionality
        /// </summary>
        DotNet,
        /// <summary>
        /// Use Excels strategy with max 15 significant figures.
        /// </summary>
        Excel
    }
}
