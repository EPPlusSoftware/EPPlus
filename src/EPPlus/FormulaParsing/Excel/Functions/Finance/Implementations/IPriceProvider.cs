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
    /// IPriceProvider
    /// </summary>
    public interface IPriceProvider
    {
        /// <summary>
        /// GetPrice
        /// </summary>
        /// <param name="settlement"></param>
        /// <param name="maturity"></param>
        /// <param name="rate"></param>
        /// <param name="yield"></param>
        /// <param name="redemption"></param>
        /// <param name="frequency"></param>
        /// <param name="basis"></param>
        /// <returns></returns>
        double GetPrice(DateTime settlement, DateTime maturity, double rate, double yield, double redemption, int frequency, DayCountBasis basis = DayCountBasis.US_30_360);
    }
}
