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
    internal class InternalMethods
    {
        internal static FinanceCalcResult PMT_Internal(double Rate, double NPer, double PV, double FV = 0, PmtDue Due = PmtDue.EndOfPeriod)
        {
            double dTemp;
            double dTemp2;
            double dTemp3;

            //       Checking for error conditions
            if (NPer == 0.0)
                return new FinanceCalcResult(eErrorType.Value);

            if(Rate == 0.0)
            {
                return new FinanceCalcResult((-FV - PV) / NPer);
            }
            else
            {
                if (Due != 0)
                    dTemp = 1.0 + Rate;
                else
                    dTemp = 1.0;
                dTemp3 = Rate + 1.0;
                //       WARSI Using the exponent operator for pow(..) in C code of PMT. Still got
                //       to make sure that they (pow and ^) are same for all conditions
                dTemp2 = System.Math.Pow(dTemp3, NPer);
                var result = ((-FV - PV * dTemp2) / (dTemp * (dTemp2 - 1.0)) * Rate);
                return new FinanceCalcResult(result);
            }
        }

        internal static double FV_Internal(double Rate, double NPer, double Pmt, double PV = 0, PmtDue Due = PmtDue.EndOfPeriod)
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
