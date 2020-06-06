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
    internal static class MirrImpl
    {
        internal static double LDoNPV(double Rate, ref double[] ValueArray, int iWNType)
        {
            bool bSkipPos;
            bool bSkipNeg;

            double dTemp2;
            double dTotal;
            double dTVal;
            int I;
            int lLower;
            int lUpper;

            bSkipPos = iWNType < 0;
            bSkipNeg = iWNType > 0;

            dTemp2 = 1.0;
            dTotal = 0.0;

            lLower = 0;
            lUpper = ValueArray.Length - 1;

            for(I = lLower; I <= lUpper; I++)
            {
                dTVal = ValueArray[I];
                dTemp2 = dTemp2 + dTemp2 * Rate;

                if(!((bSkipPos && dTVal > 0.0) || (bSkipNeg && dTVal< 0.0)))
                {
                    dTotal = dTotal + dTVal / dTemp2;
                }
           
            }
            return dTotal;
        }

        internal static FinanceCalcResult<double> MIRR(double[] ValueArray, double FinanceRate, double ReinvestRate)
        {
            double dNpvPos;
            double dNpvNeg;
            double dTemp;
            double dTemp1;
            double dNTemp2;
            int lCVal;
            int lLower;
            int lUpper;

            if(ValueArray.Rank != 1)
            {
                return new FinanceCalcResult<double>(eErrorType.Value);
            }

            lLower = 0;
            lUpper = ValueArray.Length - 1;
            lCVal = lUpper - lLower + 1;

            if(FinanceRate == -1d)
            {
                return new FinanceCalcResult<double>(eErrorType.Num);
            }

            if(ReinvestRate == -1d)
            {
                return new FinanceCalcResult<double>(eErrorType.Num);
            }

            if(lCVal <= 1d)
            {
                return new FinanceCalcResult<double>(eErrorType.Num);
            }

            dNpvNeg = LDoNPV(FinanceRate, ref ValueArray, -1);

            if (dNpvNeg == 0.0)
                return new FinanceCalcResult<double>(eErrorType.Div0);

            dNpvPos = LDoNPV(ReinvestRate, ref ValueArray, 1); // npv of +ve values
            dTemp1 = ReinvestRate + 1.0;
            dNTemp2 = lCVal;

            dTemp = -dNpvPos * System.Math.Pow(dTemp1, dNTemp2) / (dNpvNeg * (FinanceRate + 1.0));

            if (dTemp < 0d)
                return new FinanceCalcResult<double>(eErrorType.Value);

            dTemp1 = 1d / (lCVal - 1d);

            return new FinanceCalcResult<double>(System.Math.Pow(dTemp, dTemp1) - 1.0);
        }

    }
}
