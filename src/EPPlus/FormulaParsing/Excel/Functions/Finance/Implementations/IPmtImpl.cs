/*************************************************************************************************
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
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal static class IPmtImpl
    {
        internal static FinanceCalcResult Ipmt(double Rate, double Per, double NPer, double PV, double FV = 0, PmtDue Due = PmtDue.EndOfPeriod)
        {
            double Pmt;
            double dTFv;
            double dTemp;

            if(Due != PmtDue.EndOfPeriod)
            {
                dTemp = 2d;
            }
            else
            {
                dTemp = 1;
            }

            // Type = 0 or non-zero only. Offset to calculate FV
            if((Per <= 0) || (Per >= NPer + 1))
            {
                return new FinanceCalcResult(eErrorType.Value);
            }

            if(Due != PmtDue.EndOfPeriod && (Per == 1.0))
            {
                return new FinanceCalcResult(0d); ;
            }

            //   Calculate PMT (i.e. annuity) for given parms. Rqrd for FV
            var result = InternalMethods.PMT_Internal(Rate, NPer, PV, FV, Due);
            if (result.HasError) return new FinanceCalcResult(eErrorType.Num);
            Pmt = result.Result;

            if(Due != PmtDue.EndOfPeriod)
            {
                PV = PV + Pmt;
            }

            dTFv = InternalMethods.FV_Internal(Rate, (Per - dTemp), Pmt, PV, PmtDue.EndOfPeriod);

            return new FinanceCalcResult(dTFv * Rate);
        }
    }
}
