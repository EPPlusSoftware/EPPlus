using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.FinancialDayCount;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance.Implementations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Finance
{
    [TestClass]
    public class CoupInternalTests
    {
        #region COUPDAYBS
        [TestMethod]
        public void Coupdaybs_ShouldReturnCorrectResult_ActualActual()
        {
            var settlement = new DateTime(2018, 12, 01);
            var maturity = new DateTime(2019, 3, 15);

            var func = new CoupdaybsImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_Actual),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_Actual),
                4,
                DayCountBasis.Actual_Actual
                );
            var result = func.Coupdaybs();
            Assert.AreEqual(77, result.Result);

            settlement = new DateTime(2016, 02, 01);
            maturity = new DateTime(2019, 05, 31);

            func = new CoupdaybsImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_Actual),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_Actual),
                2,
                DayCountBasis.Actual_Actual
                );
            result = func.Coupdaybs();
            Assert.AreEqual(63, result.Result);
        }

        [TestMethod]
        public void Coupdaybs_ShouldReturnCorrectResult_Us_30_360()
        {
            var settlement = new DateTime(2018, 12, 01);
            var maturity = new DateTime(2019, 3, 15);

            var func = new CoupdaybsImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.US_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.US_30_360),
                4,
                DayCountBasis.US_30_360
                );
            var result = func.Coupdaybs();
            Assert.AreEqual(76, result.Result);

            settlement = new DateTime(2016, 02, 01);
            maturity = new DateTime(2019, 05, 31);

            func = new CoupdaybsImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.US_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.US_30_360),
                2,
                DayCountBasis.US_30_360
                );
            result = func.Coupdaybs();
            Assert.AreEqual(61, result.Result);
        }

        [TestMethod]
        public void Coupdaybs_ShouldReturnCorrectResult_Actual_360()
        {
            var settlement = new DateTime(2018, 12, 01);
            var maturity = new DateTime(2019, 3, 15);

            var func = new CoupdaybsImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_360),
                4,
                DayCountBasis.Actual_360
                );
            var result = func.Coupdaybs();
            Assert.AreEqual(77, result.Result);

            settlement = new DateTime(2016, 02, 01);
            maturity = new DateTime(2019, 05, 31);

            func = new CoupdaybsImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_360),
                2,
                DayCountBasis.Actual_360
                );
            result = func.Coupdaybs();
            Assert.AreEqual(63, result.Result);
        }

        [TestMethod]
        public void Coupdaybs_ShouldReturnCorrectResult_Actual_365()
        {
            var settlement = new DateTime(2018, 12, 01);
            var maturity = new DateTime(2019, 3, 15);

            var func = new CoupdaybsImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_365),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_365),
                4,
                DayCountBasis.Actual_365
                );
            var result = func.Coupdaybs();
            Assert.AreEqual(77, result.Result);

            settlement = new DateTime(2016, 02, 01);
            maturity = new DateTime(2019, 05, 31);

            func = new CoupdaybsImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_365),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_365),
                2,
                DayCountBasis.Actual_365
                );
            result = func.Coupdaybs();
            Assert.AreEqual(63, result.Result);
        }

        [TestMethod]
        public void Coupdaybs_ShouldReturnCorrectResult_European_30_360()
        {
            var settlement = new DateTime(2018, 12, 01);
            var maturity = new DateTime(2019, 3, 15);

            var func = new CoupdaybsImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.European_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.European_30_360),
                4,
                DayCountBasis.European_30_360
                );
            var result = func.Coupdaybs();
            Assert.AreEqual(76, result.Result);

            settlement = new DateTime(2016, 02, 01);
            maturity = new DateTime(2019, 05, 31);

            func = new CoupdaybsImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.European_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.European_30_360),
                2,
                DayCountBasis.European_30_360
                );
            result = func.Coupdaybs();
            Assert.AreEqual(61, result.Result);
        }

        #endregion

        #region COUPDAYS
        [TestMethod]
        public void Coupdays_ShouldReturnNumberOfDays_ActualActual()
        {
            var settlement = new DateTime(2012, 2, 29);
            var maturity = new DateTime(2019, 3, 15);

            var func = new CoupdaysImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_Actual),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_Actual),
                4,
                DayCountBasis.Actual_Actual
                );
            var result = func.GetCoupdays();
            Assert.AreEqual(91d, result.Result);

            //settlement = new DateTime(2016, 2, 1);
            //maturity = new DateTime(2019, 5, 31);
            settlement = new DateTime(2017, 2, 1);
            maturity = new DateTime(2019, 5, 31);

            func = new CoupdaysImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_Actual),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_Actual),
                4,
                DayCountBasis.Actual_Actual
                );
            result = func.GetCoupdays();
            Assert.AreEqual(90d, result.Result);
        }

        [TestMethod]
        public void Coupdays_ShouldReturnNumberOfDays_Us_30_360()
        {
            var settlement = new DateTime(2012, 2, 29);
            var maturity = new DateTime(2019, 3, 15);

            var func = new CoupdaysImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.US_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.US_30_360),
                4,
                DayCountBasis.US_30_360
                );
            var result = func.GetCoupdays();
            Assert.AreEqual(90d, result.Result);


            settlement = new DateTime(2017, 2, 1);
            maturity = new DateTime(2019, 5, 31);

            func = new CoupdaysImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.US_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.US_30_360),
                4,
                DayCountBasis.US_30_360
                );
            result = func.GetCoupdays();
            Assert.AreEqual(90d, result.Result);
        }

        [TestMethod]
        public void Coupdays_ShouldReturnNumberOfDays_Actual_360()
        {
            var settlement = new DateTime(2012, 2, 29);
            var maturity = new DateTime(2019, 3, 15);

            var func = new CoupdaysImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_360),
                4,
                DayCountBasis.Actual_360
                );
            var result = func.GetCoupdays();
            Assert.AreEqual(90d, result.Result);


            settlement = new DateTime(2017, 2, 1);
            maturity = new DateTime(2019, 5, 31);

            func = new CoupdaysImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_360),
                4,
                DayCountBasis.Actual_360
                );
            result = func.GetCoupdays();
            Assert.AreEqual(90d, result.Result);
        }

        [TestMethod]
        public void Coupdays_ShouldReturnNumberOfDays_Actual_365()
        {
            var settlement = new DateTime(2012, 2, 29);
            var maturity = new DateTime(2019, 3, 15);

            var func = new CoupdaysImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_365),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_365),
                4,
                DayCountBasis.Actual_365
                );
            var result = func.GetCoupdays();
            Assert.AreEqual(91.25, result.Result);


            settlement = new DateTime(2017, 2, 1);
            maturity = new DateTime(2019, 5, 31);

            func = new CoupdaysImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_365),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_365),
                4,
                DayCountBasis.Actual_365
                );
            result = func.GetCoupdays();
            Assert.AreEqual(91.25, result.Result);
        }

        [TestMethod]
        public void Coupdays_ShouldReturnNumberOfDays_European_Actual_360()
        {
            var settlement = new DateTime(2012, 2, 29);
            var maturity = new DateTime(2019, 3, 15);

            var func = new CoupdaysImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.European_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.European_30_360),
                4,
                DayCountBasis.European_30_360
                );
            var result = func.GetCoupdays();
            Assert.AreEqual(90d, result.Result);


            settlement = new DateTime(2017, 2, 1);
            maturity = new DateTime(2019, 5, 31);

            func = new CoupdaysImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.European_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.European_30_360),
                4,
                DayCountBasis.European_30_360
                );
            result = func.GetCoupdays();
            Assert.AreEqual(90d, result.Result);
        }

        #endregion

        #region COUPDAYSNC
        [TestMethod]
        public void CoupdaysNc_ShouldReturnNumberOfDays_ActualActual()
        {
            var settlement = new DateTime(2016, 02, 01);
            var maturity = new DateTime(2019, 5, 31);

            var func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_Actual),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_Actual),
                4,
                DayCountBasis.Actual_Actual
                );
            var result = func.Coupdaysnc();
            Assert.AreEqual(28d, result.Result);

            settlement = new DateTime(2016, 05, 30);
            maturity = new DateTime(2019, 5, 31);

            func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_Actual),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_Actual),
                4,
                DayCountBasis.Actual_Actual
                );
            result = func.Coupdaysnc();
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void CoupdaysNc_ShouldReturnNumberOfDays_Us_30_360()
        {
            var settlement = new DateTime(2016, 02, 01);
            var maturity = new DateTime(2019, 5, 31);

            var func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.US_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.US_30_360),
                4,
                DayCountBasis.US_30_360
                );
            var result = func.Coupdaysnc();
            Assert.AreEqual(29d, result.Result);

            settlement = new DateTime(2016, 05, 30);
            maturity = new DateTime(2019, 5, 31);

            func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.US_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.US_30_360),
                4,
                DayCountBasis.US_30_360
                );
            result = func.Coupdaysnc();
            Assert.AreEqual(0d, result.Result);

            settlement = new DateTime(2018, 08, 01);
            maturity = new DateTime(2019, 03, 15);

            func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.US_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.US_30_360),
                4,
                DayCountBasis.US_30_360
                );
            result = func.Coupdaysnc();
            Assert.AreEqual(44d, result.Result);

            // NB! Excel returns 29 on this one. Google sheets returns 28. As far as we can see it should be
            // 28, so that's how we have implemented it in EPPlus
            settlement = new DateTime(2018, 08, 01);
            maturity = new DateTime(2020, 02, 29);

            func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.US_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.US_30_360),
                4,
                DayCountBasis.US_30_360
                );
            result = func.Coupdaysnc();
            Assert.AreEqual(28d, result.Result);
        }

        [TestMethod]
        public void CoupdaysNc_ShouldReturnNumberOfDays_Actual_360()
        {
            var settlement = new DateTime(2016, 02, 01);
            var maturity = new DateTime(2019, 5, 31);

            var func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_360),
                4,
                DayCountBasis.Actual_360
                );
            var result = func.Coupdaysnc();
            Assert.AreEqual(28d, result.Result);

            settlement = new DateTime(2016, 05, 30);
            maturity = new DateTime(2019, 5, 31);

            func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_360),
                4,
                DayCountBasis.Actual_360
                );
            result = func.Coupdaysnc();
            Assert.AreEqual(1d, result.Result);

            settlement = new DateTime(2018, 08, 01);
            maturity = new DateTime(2019, 03, 15);

            func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_360),
                4,
                DayCountBasis.Actual_360
                );
            result = func.Coupdaysnc();
            Assert.AreEqual(45d, result.Result);
        }

        [TestMethod]
        public void CoupdaysNc_ShouldReturnNumberOfDays_Actual_365()
        {
            var settlement = new DateTime(2016, 02, 01);
            var maturity = new DateTime(2019, 5, 31);

            var func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_365),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_365),
                4,
                DayCountBasis.Actual_365
                );
            var result = func.Coupdaysnc();
            Assert.AreEqual(28d, result.Result);

            settlement = new DateTime(2016, 05, 30);
            maturity = new DateTime(2019, 5, 31);

            func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_365),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_365),
                4,
                DayCountBasis.Actual_365
                );
            result = func.Coupdaysnc();
            Assert.AreEqual(1d, result.Result);

            settlement = new DateTime(2018, 08, 01);
            maturity = new DateTime(2019, 03, 15);

            func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_365),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_365),
                4,
                DayCountBasis.Actual_365
                );
            result = func.Coupdaysnc();
            Assert.AreEqual(45d, result.Result);
        }

        [TestMethod]
        public void CoupdaysNc_ShouldReturnNumberOfDays_European_30_360()
        {
            var settlement = new DateTime(2016, 02, 01);
            var maturity = new DateTime(2019, 5, 31);

            var func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.European_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.European_30_360),
                4,
                DayCountBasis.European_30_360
                );
            var result = func.Coupdaysnc();
            Assert.AreEqual(28d, result.Result);

            settlement = new DateTime(2016, 05, 30);
            maturity = new DateTime(2019, 5, 31);

            func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.European_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.European_30_360),
                4,
                DayCountBasis.European_30_360
                );
            result = func.Coupdaysnc();
            Assert.AreEqual(0d, result.Result);

            settlement = new DateTime(2018, 08, 01);
            maturity = new DateTime(2019, 03, 15);

            func = new CoupdaysncImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.European_30_360),
                FinancialDayFactory.Create(maturity, DayCountBasis.European_30_360),
                4,
                DayCountBasis.European_30_360
                );
            result = func.Coupdaysnc();
            Assert.AreEqual(44d, result.Result);
        }

        #endregion

        #region COUPNUM

        [TestMethod]
        public void Coupnum_ShouldReturnNumberOfDays_ActualActual()
        {
            // No need for tests per DayCountBasis here.
            var settlement = new DateTime(2016, 02, 01);
            var maturity = new DateTime(2019, 3, 15);

            var func = new CoupnumImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_Actual),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_Actual),
                4,
                DayCountBasis.Actual_Actual
                );
            var result = func.GetCoupnum();
            Assert.AreEqual(13d, result.Result);

            settlement = new DateTime(2018, 12, 01);
            maturity = new DateTime(2019, 5, 31);

            func = new CoupnumImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_Actual),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_Actual),
                2,
                DayCountBasis.Actual_Actual
                );
            result = func.GetCoupnum();
            Assert.AreEqual(1d, result.Result);
        }
        #endregion

        #region COUPNCD
        [TestMethod]
        public void Coupncd_ShouldReturnCorrectDate_ActualActual()
        {
            // No need for tests per DayCountBasis here.
            var settlement = new DateTime(2017, 02, 01);
            var maturity = new DateTime(2020, 05, 31);

            var func = new CoupncdImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_Actual),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_Actual),
                4,
                DayCountBasis.Actual_Actual
                );
            var result = func.GetCoupncd();
            Assert.AreEqual(new System.DateTime(2017, 2, 28), result.Result);

            settlement = new DateTime(2016, 02, 01);
            maturity = new DateTime(2019, 5, 31);

            func = new CoupncdImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_Actual),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_Actual),
                4,
                DayCountBasis.Actual_Actual
                );
            result = func.GetCoupncd();
            Assert.AreEqual(new System.DateTime(2016, 2, 29), result.Result);
        }
        #endregion

        #region COUPPCD
        [TestMethod]
        public void Couppcd_ShouldReturnCorrectDate_ActualActual()
        {
            // No need for tests per DayCountBasis here.
            var settlement = new DateTime(2017, 05, 30);
            var maturity = new DateTime(2020, 05, 31);

            var func = new CouppcdImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_Actual),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_Actual),
                4,
                DayCountBasis.Actual_Actual
                );
            var result = func.GetCouppcd();
            Assert.AreEqual(new System.DateTime(2017, 2, 28), result.Result);

            settlement = new DateTime(2016, 02, 01);
            maturity = new DateTime(2019, 5, 31);

            func = new CouppcdImpl(
                FinancialDayFactory.Create(settlement, DayCountBasis.Actual_Actual),
                FinancialDayFactory.Create(maturity, DayCountBasis.Actual_Actual),
                4,
                DayCountBasis.Actual_Actual
                );
            result = func.GetCouppcd();
            Assert.AreEqual(new System.DateTime(2015, 11, 30), result.Result);
        }
        #endregion
    }
}
