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
    internal class FinancialPeriod
    {
        public FinancialPeriod(FinancialDay start, FinancialDay end)
        {
            Start = start;
            End = end;
        }
        internal FinancialDay Start { get; }

        internal FinancialDay End { get; }

        public override string ToString()
        {
            return $"{Start.ToString()} - {End.ToString()}";
        }
    }
}
