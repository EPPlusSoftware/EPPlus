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
    internal class CumprincImpl
    {
        public CumprincImpl(IPmtProvider pmtProvider, IFvProvider fvProvider)
        {
            _pmtProvider = pmtProvider;
            _fvProvider = fvProvider;
        }

        private readonly IPmtProvider _pmtProvider;
        private readonly IFvProvider _fvProvider;

        public FinanceCalcResult<double> GetCumprinc(double rate, double nper, double pv, int startPeriod, int endPeriod, PmtDue type)
        {
            double fPmt, fPpmt;

            if (startPeriod < 1 || endPeriod < startPeriod || rate <= 0.0 || endPeriod > nper || pv <= 0.0)
                return new FinanceCalcResult<double>(eErrorType.Num);

            fPmt = _pmtProvider.GetPmt(rate, nper, pv, 0.0, type);

            fPpmt = 0.0;

            var nStart = startPeriod;
            var nEnd = endPeriod;

            if (nStart == 1)
            {
                if (type == PmtDue.EndOfPeriod)
                    fPpmt = fPmt + pv * rate;
                else
                    fPpmt = fPmt;

                nStart++;
            }

            for (var i = nStart; i <= nEnd; i++)
            {
                if (type == PmtDue.BeginningOfPeriod)
                    fPpmt += fPmt - (_fvProvider.GetFv(rate, i - 2, fPmt, pv, type) - fPmt) * rate;
                else
                    fPpmt += fPmt - _fvProvider.GetFv(rate, i - 1, fPmt, pv, type) * rate;
            }

            return new FinanceCalcResult<double>(fPpmt);
        }
    }
}
