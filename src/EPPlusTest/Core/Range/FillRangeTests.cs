using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;

namespace EPPlusTest.Core.Range
{
    [TestClass]
    public class RangeFillTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _wsNum;
        static ExcelWorksheet _wsDate;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Range_Fill.xlsx", true);
            _wsNum = _pck.Workbook.Worksheets.Add("FillNumbers");
            _wsDate = _pck.Workbook.Worksheets.Add("FillDates");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }

        [TestMethod]
        public void FillNumbers_WithStartAndStep()
        {
            _wsNum.Cells["A1:B5"].FillNumbers(2, 3);
            //Assert
            Assert.AreEqual(3D, _wsNum.Cells["A1"].Value);
            Assert.AreEqual(5D, _wsNum.Cells["A2"].Value);
            Assert.AreEqual(7D, _wsNum.Cells["A3"].Value);
            Assert.AreEqual(9D, _wsNum.Cells["A4"].Value);
            Assert.AreEqual(11D, _wsNum.Cells["A5"].Value);

            Assert.AreEqual(3D, _wsNum.Cells["B1"].Value);
            Assert.AreEqual(11D, _wsNum.Cells["B5"].Value);
        }
        [TestMethod]
        public void FillNumbers()
        {
            _wsNum.Cells["C1"].Value = 7D;
            _wsNum.Cells["D1"].Value = "D1";
            _wsNum.Cells["D2"].Value = "D2";
            _wsNum.Cells["C1:D5"].FillNumbers();
            //Assert
            Assert.AreEqual(7D, _wsNum.Cells["C1"].Value);
            Assert.AreEqual(8D, _wsNum.Cells["C2"].Value);
            Assert.AreEqual(9D, _wsNum.Cells["C3"].Value);
            Assert.AreEqual(10D, _wsNum.Cells["C4"].Value);
            Assert.AreEqual(11D, _wsNum.Cells["C5"].Value);

            Assert.AreEqual("D1", _wsNum.Cells["D1"].Value);
            Assert.IsNull(_wsNum.Cells["D2"].Value);
        }
        [TestMethod]
        public void FillNumbers_RowWithStartAndStep()
        {
            _wsNum.Cells["A10:E11"].FillNumbers(x =>
            {
                x.Direction = eFillDirection.Row;
                x.StartValue = 3;
                x.StepValue = 2;
            }); 
            //Assert
            Assert.AreEqual(3D, _wsNum.Cells["A10"].Value);
            Assert.AreEqual(5D, _wsNum.Cells["B10"].Value);
            Assert.AreEqual(7D, _wsNum.Cells["C10"].Value);
            Assert.AreEqual(9D, _wsNum.Cells["D10"].Value);
            Assert.AreEqual(11D, _wsNum.Cells["E10"].Value);

            Assert.AreEqual(3D, _wsNum.Cells["A11"].Value);
            Assert.AreEqual(11D, _wsNum.Cells["E11"].Value);
        }
        [TestMethod]
        public void FillNumbers_RowWithStartAndStepMultipy()
        {
            _wsNum.Cells["A13:E14"].FillNumbers(x =>
            {
                x.CalculationMethod = eCalculationMethod.Multiply;
                x.Direction = eFillDirection.Row;
                x.StartValue = 3;
                x.StepValue = 2;
            });
            //Assert
            Assert.AreEqual(3D, _wsNum.Cells["A10"].Value);
            Assert.AreEqual(6D, _wsNum.Cells["B10"].Value);
            Assert.AreEqual(12D, _wsNum.Cells["C10"].Value);
            Assert.AreEqual(24D, _wsNum.Cells["D10"].Value);
            Assert.AreEqual(48D, _wsNum.Cells["E10"].Value);

            Assert.AreEqual(3D, _wsNum.Cells["A11"].Value);
            Assert.AreEqual(48D, _wsNum.Cells["E11"].Value);
        }
        [TestMethod]
        public void FillDate_WithStartAndStepDay()
        {
            var startDate = new DateTime(2021, 1, 1);
            _wsNum.Cells["A1:B5"].FillDates(new DateTime(2021,1,1));
            //Assert
            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsNum.Cells["A1"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay, ((DateTime)_wsNum.Cells["A2"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay*2, ((DateTime)_wsNum.Cells["A3"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay*3, ((DateTime)_wsNum.Cells["A4"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay*4, ((DateTime)_wsNum.Cells["A5"].Value).Ticks);

            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay, ((DateTime)_wsNum.Cells["B2"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 4, ((DateTime)_wsNum.Cells["B5"].Value).Ticks);
        }
        [TestMethod]
        public void FillDate()
        {
            var startDate = new DateTime(2021, 6, 30);
            _wsDate.Cells["C1"].Value = startDate;
            _wsDate.Cells["D1"].Value = "D1";
            _wsDate.Cells["C1:D5"].FillDates();
            //Assert
            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsDate.Cells["C1"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay, ((DateTime)_wsDate.Cells["C2"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 2, ((DateTime)_wsDate.Cells["C3"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 3, ((DateTime)_wsDate.Cells["C4"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 4, ((DateTime)_wsDate.Cells["C5"].Value).Ticks);

            Assert.AreEqual("D1", _wsDate.Cells["D1"].Value);
            Assert.IsNull(_wsDate.Cells["D2"].Value);
        }

    }
}
