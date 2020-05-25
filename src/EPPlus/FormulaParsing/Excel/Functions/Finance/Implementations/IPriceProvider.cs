using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    public interface IPriceProvider
    {
        double GetPrice(System.DateTime settlement, System.DateTime maturity, double rate, double yield, double redemption, int frequency, DayCountBasis basis = DayCountBasis.US_30_360);
    }
}
