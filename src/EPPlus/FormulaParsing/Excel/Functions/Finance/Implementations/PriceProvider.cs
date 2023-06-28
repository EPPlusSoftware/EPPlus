﻿/*************************************************************************************************
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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal class PriceProvider : IPriceProvider
    {
        public double GetPrice(DateTime settlement, DateTime maturity, double rate, double yield, double redemption, int frequency, DayCountBasis basis = DayCountBasis.US_30_360)
        {
            return PriceImpl.GetPrice(
                FinancialDayFactory.Create(settlement, basis),
                FinancialDayFactory.Create(maturity, basis),
                rate,
                yield,
                redemption,
                frequency,
                basis).Result;
        }
    }
}
