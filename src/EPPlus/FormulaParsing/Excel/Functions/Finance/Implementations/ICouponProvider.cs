using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    public interface ICouponProvider
    {
        double GetCoupdaybs(System.DateTime settlement, System.DateTime maturity, int frequency, DayCountBasis basis);

        double GetCoupdays(System.DateTime settlement, System.DateTime maturity, int frequency, DayCountBasis basis);

        double GetCoupdaysnc(System.DateTime settlement, System.DateTime maturity, int frequency, DayCountBasis basis);

        System.DateTime GetCoupsncd(System.DateTime settlement, System.DateTime maturity, int frequency, DayCountBasis basis);

        double GetCoupnum(System.DateTime settlement, System.DateTime maturity, int frequency, DayCountBasis basis);

        System.DateTime GetCouppcd(System.DateTime settlement, System.DateTime maturity, int frequency, DayCountBasis basis);
    }
}
