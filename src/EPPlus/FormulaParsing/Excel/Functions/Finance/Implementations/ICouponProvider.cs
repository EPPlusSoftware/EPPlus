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
    public interface ICouponProvider
    {
        double GetCoupdaybs(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis);

        double GetCoupdays(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis);

        double GetCoupdaysnc(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis);

        DateTime GetCoupsncd(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis);

        double GetCoupnum(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis);

        DateTime GetCouppcd(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis);
    }
}
