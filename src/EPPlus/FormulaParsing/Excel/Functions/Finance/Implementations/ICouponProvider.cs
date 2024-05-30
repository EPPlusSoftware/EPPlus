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
    /// ICouponProvider
    /// </summary>
    public interface ICouponProvider
    {
        /// <summary>
        /// GetCoupdaybs
        /// </summary>
        /// <param name="settlement"></param>
        /// <param name="maturity"></param>
        /// <param name="frequency"></param>
        /// <param name="basis"></param>
        /// <returns></returns>
        double GetCoupdaybs(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis);
        /// <summary>
        /// CoupDays
        /// </summary>
        /// <param name="settlement"></param>
        /// <param name="maturity"></param>
        /// <param name="frequency"></param>
        /// <param name="basis"></param>
        /// <returns></returns>
        double GetCoupdays(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis);
        /// <summary>
        /// Coupdaysnc
        /// </summary>
        /// <param name="settlement"></param>
        /// <param name="maturity"></param>
        /// <param name="frequency"></param>
        /// <param name="basis"></param>
        /// <returns></returns>
        double GetCoupdaysnc(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis);
        /// <summary>
        /// GetCoupsncd
        /// </summary>
        /// <param name="settlement"></param>
        /// <param name="maturity"></param>
        /// <param name="frequency"></param>
        /// <param name="basis"></param>
        /// <returns></returns>
        DateTime GetCoupsncd(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis);
        /// <summary>
        /// GetCoupnum
        /// </summary>
        /// <param name="settlement"></param>
        /// <param name="maturity"></param>
        /// <param name="frequency"></param>
        /// <param name="basis"></param>
        /// <returns></returns>
        double GetCoupnum(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis);
        /// <summary>
        /// GetCouppcd
        /// </summary>
        /// <param name="settlement"></param>
        /// <param name="maturity"></param>
        /// <param name="frequency"></param>
        /// <param name="basis"></param>
        /// <returns></returns>
        DateTime GetCouppcd(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis);
    }
}
