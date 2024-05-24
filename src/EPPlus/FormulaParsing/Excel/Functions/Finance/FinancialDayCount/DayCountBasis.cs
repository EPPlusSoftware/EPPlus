/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount
{
    /// <summary>
    /// Day counting options unsed in the internal finance function implementations.
    /// </summary>
    public enum DayCountBasis
    {
        /// <summary>
        /// US basis
        /// </summary>
        US_30_360 = 0,
        /// <summary>
        /// Actual actual
        /// </summary>
        Actual_Actual = 1,
        /// <summary>
        /// Actual 360
        /// </summary>
        Actual_360 = 2,
        /// <summary>
        /// Actual 365
        /// </summary>
        Actual_365 = 3,
        /// <summary>
        /// Europe
        /// </summary>
        European_30_360 = 4
    }
}
