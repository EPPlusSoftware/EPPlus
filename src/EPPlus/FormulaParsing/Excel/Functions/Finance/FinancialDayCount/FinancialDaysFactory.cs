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
    internal static class FinancialDaysFactory
    {
        internal static IFinanicalDays Create(DayCountBasis basis)
        {
            switch(basis)
            {
                case DayCountBasis.US_30_360:
                    return new FinancialDaysUs_30_360();
                case DayCountBasis.Actual_Actual:
                    return new FinancialDays_Actual_Actual();
                case DayCountBasis.Actual_360:
                    return new FinancialDays_Actual_360();
                case DayCountBasis.Actual_365:
                    return new FinancialDays_Actual_365();
                case DayCountBasis.European_30_360:
                    return new FinancialDaysEuropean_30_360();
                default:
                    throw new ArgumentException("basis");
            }
        }
    }
}
