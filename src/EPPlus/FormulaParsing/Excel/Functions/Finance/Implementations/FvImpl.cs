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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal static class FvImpl
    {
        internal static FinanceCalcResult<double> Fv(double Rate, double NPer, double Pmt, double PV = 0, PmtDue Due = PmtDue.EndOfPeriod)
        {
            return new FinanceCalcResult<double>(InternalMethods.FV_Internal(Rate, NPer, Pmt, PV, Due));
        }
    }
}
