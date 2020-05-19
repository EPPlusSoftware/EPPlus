/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/03/2020         EPPlus Software AB         Implemented function (ported to c# from Microsoft.VisualBasic.Financial.vb (MIT))
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal static class PpmtImpl
    {
        internal static FinanceCalcResult<double> Ppmt(double Rate, double Per, double NPer, double PV, double FV = 0, PmtDue Due = PmtDue.EndOfPeriod)
        {
            double Pmt;
            double dIPMT;

            //   Checking for error conditions
            if ((Per <= 0.0) || (Per >= (NPer + 1)))
                return new FinanceCalcResult<double>(eErrorType.Num);

            var pmtResult = InternalMethods.PMT_Internal(Rate, NPer, PV, FV, Due);
            if (pmtResult.HasError) return new FinanceCalcResult<double>(pmtResult.ExcelErrorType);
            Pmt = pmtResult.Result;

            var iPmtResult = IPmtImpl.Ipmt(Rate, Per, NPer, PV, FV, Due);
            if (iPmtResult.HasError) return new FinanceCalcResult<double>(iPmtResult.ExcelErrorType);
            dIPMT = iPmtResult.Result;

            return new FinanceCalcResult<double>(Pmt - dIPMT);
        }
    }
}
