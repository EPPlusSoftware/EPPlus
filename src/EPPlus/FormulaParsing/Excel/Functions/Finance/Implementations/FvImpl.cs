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
        internal static FinanceCalcResult Fv(double Rate, double NPer, double Pmt, double PV = 0, PmtDue Due = PmtDue.EndOfPeriod)
        {
            return new FinanceCalcResult(FV_Internal(Rate, NPer, Pmt, PV, Due));
        }

        private static double FV_Internal(double Rate, double NPer, double Pmt, double PV = 0, PmtDue Due = PmtDue.EndOfPeriod)
        {
            double dTemp;
            double dTemp2;
            double dTemp3;

            //Performing calculation
            if (Rate == 0)
                return (-PV - Pmt * NPer);
            if (Due != PmtDue.EndOfPeriod)
            {
                dTemp = 1.0 + Rate;
            }
            else
            {
                dTemp = 1.0;
            }

            dTemp3 = 1.0 + Rate;
            dTemp2 = System.Math.Pow(dTemp3, NPer);

            //Do divides before multiplies to avoid OverflowExceptions
            return ((-PV) * dTemp2) - ((Pmt / Rate) * dTemp * (dTemp2 - 1.0));
        }
    }
}
