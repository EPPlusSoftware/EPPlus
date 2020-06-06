using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal static class CumipmtImpl
    {
        public static FinanceCalcResult<double> GetCumipmt(double rate, int nPer, double pv, int startPeriod, int endPeriod, PmtDue type )
        {
            if (startPeriod <= 0 || endPeriod < startPeriod || rate <= 0d || endPeriod > nPer ||
                pv <= 0d)
                return new FinanceCalcResult<double>(eErrorType.Num);

            
            var result = InternalMethods.PMT_Internal(rate, nPer, pv, 0d, type );
            if (result.HasError) return new FinanceCalcResult<double>(result.ExcelErrorType);
            var pmtResult = result.Result;
      
            var retVal = 0d;
      
            if(startPeriod == 1 )
            {
                if(type == PmtDue.EndOfPeriod )
                    retVal = -pv;

                startPeriod++;
            }
      
            for(int i = startPeriod; i <= endPeriod; i++ )
            {
                var res = FvImpl.Fv(rate, (i - 1 - (int)type), pmtResult, pv, type);
                if (res.HasError) return new FinanceCalcResult<double>(res.ExcelErrorType);
                retVal += type == PmtDue.BeginningOfPeriod ? res.Result - pmtResult : res.Result;
            }
      
            retVal *= rate;
      
            return new FinanceCalcResult<double>(retVal);
        }
    }
}
