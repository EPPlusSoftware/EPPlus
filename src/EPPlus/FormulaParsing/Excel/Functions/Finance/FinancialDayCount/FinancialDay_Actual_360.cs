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
    internal class FinancialDay_Actual_360 : FinancialDay
    {
        public FinancialDay_Actual_360(DateTime date) : base(date)
        {
        }

        public FinancialDay_Actual_360(int year, int month, int day) : base(year, month, day)
        {
        }

        protected override DayCountBasis Basis => DayCountBasis.Actual_360;

        protected override FinancialDay Factory(short year, short month, short day)
        {
            return new FinancialDay_Actual_360(year, month, day);
        }
    }
}
