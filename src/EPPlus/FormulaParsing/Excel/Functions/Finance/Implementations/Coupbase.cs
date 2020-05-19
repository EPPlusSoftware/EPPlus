/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  13/5/2020         EPPlus Software AB       Implemented Excel COUP functions
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal abstract class Coupbase
    {
        public Coupbase(FinancialDay settlement, FinancialDay maturity, int frequency, DayCountBasis basis)
        {
            Settlement = settlement;
            Maturity = maturity;
            Frequency = frequency;
            Basis = basis;
        }

        protected FinancialDay Settlement { get; }
        protected FinancialDay Maturity { get; }
        protected int Frequency { get; }
        protected DayCountBasis Basis { get; }

        protected FinancialDay GetCouponPeriodBySettlement()
        {
            var financialDays = FinancialDaysFactory.Create(Basis);
            throw new NotImplementedException();
        }
    }
}
