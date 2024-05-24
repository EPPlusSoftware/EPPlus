/*************************************************************************************************
  * This Source Code Form is subject to the terms of the Mozilla Public
  * License, v. 2.0. If a copy of the MPL was not distributed with this
  * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/20/2020         EPPlus Software AB       Implemented function
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    /// <summary>
    /// IYearFracProvider
    /// </summary>
    public interface IYearFracProvider
    {
        /// <summary>
        /// GetYearFrac
        /// </summary>
        /// <param name="date1"></param>
        /// <param name="date2"></param>
        /// <param name="basis"></param>
        /// <returns></returns>
        double GetYearFrac(DateTime date1, DateTime date2, DayCountBasis basis);
    }
}
