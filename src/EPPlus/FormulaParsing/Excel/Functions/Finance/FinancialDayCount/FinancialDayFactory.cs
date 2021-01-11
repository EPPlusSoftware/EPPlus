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
    internal static class FinancialDayFactory
    {
        internal static FinancialDay Create(System.DateTime date, DayCountBasis basis)
        {
            switch (basis)
            {
                case DayCountBasis.US_30_360:
                    return new FinancialDay_Us_30_360(date);
                case DayCountBasis.Actual_Actual:
                    return new FinancialDay_Actual_Actual(date);
                case DayCountBasis.Actual_360:
                    return new FinancialDay_Actual_360(date);
                case DayCountBasis.Actual_365:
                    return new FinancialDay_Actual_365(date);
                case DayCountBasis.European_30_360:
                    return new FinancialDay_European_30_360(date);
                default:
                    throw new ArgumentException("basis");
            }
        }

        internal static FinancialPeriod CreatePeriod(System.DateTime start, System.DateTime end, DayCountBasis basis)
        {
            var s = Create(start, basis);
            var e = Create(end, basis);
            return new FinancialPeriod(s, e);
        }
    }
}
