using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations
{
    internal class CouponProvider : ICouponProvider
    {
        public double GetCoupdaybs(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis)
        {
            var func = new CoupdaybsImpl(FinancialDayFactory.Create(settlement, basis), FinancialDayFactory.Create(maturity, basis), frequency, basis);
            return func.Coupdaybs().Result;
        }

        public double GetCoupdays(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis)
        {
            var func = new CoupdaysImpl(FinancialDayFactory.Create(settlement, basis), FinancialDayFactory.Create(maturity, basis), frequency, basis);
            return func.GetCoupdays().Result;
        }

        public double GetCoupdaysnc(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis)
        {
            var func = new CoupdaysncImpl(FinancialDayFactory.Create(settlement, basis), FinancialDayFactory.Create(maturity, basis), frequency, basis);
            return func.Coupdaysnc().Result;
        }

        public double GetCoupnum(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis)
        {
            var func = new CoupnumImpl(FinancialDayFactory.Create(settlement, basis), FinancialDayFactory.Create(maturity, basis), frequency, basis);
            return func.GetCoupnum().Result;
        }

        public DateTime GetCouppcd(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis)
        {
            var func = new CouppcdImpl(FinancialDayFactory.Create(settlement, basis), FinancialDayFactory.Create(maturity, basis), frequency, basis);
            return func.GetCouppcd().Result;
        }

        public DateTime GetCoupsncd(DateTime settlement, DateTime maturity, int frequency, DayCountBasis basis)
        {
            var func = new CoupncdImpl(FinancialDayFactory.Create(settlement, basis), FinancialDayFactory.Create(maturity, basis), frequency, basis);
            return func.GetCoupncd().Result;
        }
    }
}
