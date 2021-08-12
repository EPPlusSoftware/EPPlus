using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;

namespace EPPlusTest.Core.Range
{
    [TestClass]
    public class RangeFillTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Range_Fill.xlsx", true);
            _ws = _pck.Workbook.Worksheets.Add("FillNumbers");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }

        [TestMethod]
        public void FillNumbers_WithStartAndStep()
        {
            _ws.Cells["A1:B5"].FillNumbers(2, 3);
            //Assert
            Assert.AreEqual(3D, _ws.Cells["A1"].Value);
            Assert.AreEqual(5D, _ws.Cells["A2"].Value);
            Assert.AreEqual(7D, _ws.Cells["A3"].Value);
            Assert.AreEqual(9D, _ws.Cells["A4"].Value);
            Assert.AreEqual(11D, _ws.Cells["A5"].Value);

            Assert.AreEqual(3D, _ws.Cells["B1"].Value);
            Assert.AreEqual(11D, _ws.Cells["B5"].Value);
        }
        [TestMethod]
        public void FillNumbers_WithStep()
        {
            _ws.Cells["C1"].Value = 7D;
            _ws.Cells["D1"].Value = "D1";
            _ws.Cells["D2"].Value = "D2";
            _ws.Cells["C1:D5"].FillNumbers(1);
            //Assert
            Assert.AreEqual(7D, _ws.Cells["C1"].Value);
            Assert.AreEqual(8D, _ws.Cells["C2"].Value);
            Assert.AreEqual(9D, _ws.Cells["C3"].Value);
            Assert.AreEqual(10D, _ws.Cells["C4"].Value);
            Assert.AreEqual(11D, _ws.Cells["C5"].Value);

            Assert.AreEqual("D1", _ws.Cells["D1"].Value);
            Assert.IsNull(_ws.Cells["D2"].Value);
        }
        [TestMethod]
        public void FillNumbers_RowWithStartAndStep()
        {
            _ws.Cells["A10:E11"].FillNumbers(x =>
            {
                x.Direction = eFillDirection.Row;
                x.StartValue = 3;
                x.StepValue = 2;
            }); 
            //Assert
            Assert.AreEqual(3D, _ws.Cells["A10"].Value);
            Assert.AreEqual(5D, _ws.Cells["B10"].Value);
            Assert.AreEqual(7D, _ws.Cells["C10"].Value);
            Assert.AreEqual(9D, _ws.Cells["D10"].Value);
            Assert.AreEqual(11D, _ws.Cells["E10"].Value);

            Assert.AreEqual(3D, _ws.Cells["A11"].Value);
            Assert.AreEqual(11D, _ws.Cells["E11"].Value);
        }
        [TestMethod]
        public void FillNumbers_RowWithStartAndStepMultipy()
        {
            _ws.Cells["A13:E14"].FillNumbers(x =>
            {
                x.CalculationMethod = eCalculationMethod.Multiply;
                x.Direction = eFillDirection.Row;
                x.StartValue = 3;
                x.StepValue = 2;
            });
            //Assert
            Assert.AreEqual(3D, _ws.Cells["A10"].Value);
            Assert.AreEqual(6D, _ws.Cells["B10"].Value);
            Assert.AreEqual(12D, _ws.Cells["C10"].Value);
            Assert.AreEqual(24D, _ws.Cells["D10"].Value);
            Assert.AreEqual(48D, _ws.Cells["E10"].Value);

            Assert.AreEqual(3D, _ws.Cells["A11"].Value);
            Assert.AreEqual(48D, _ws.Cells["E11"].Value);
        }
        [TestMethod]
        public void FillDate_WithStartAndStep()
        {
            var startDate = new DateTime(2021, 1, 1);
            _ws.Cells["A1:B5"].FillDates(new DateTime(2021,1,1));
            //Assert
            Assert.AreEqual(startDate.Ticks, ((DateTime)_ws.Cells["A1"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay, ((DateTime)_ws.Cells["A2"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay*2, ((DateTime)_ws.Cells["A3"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay*3, ((DateTime)_ws.Cells["A4"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay*4, ((DateTime)_ws.Cells["A5"].Value).Ticks);

            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay, ((DateTime)_ws.Cells["B2"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 4, ((DateTime)_ws.Cells["B5"].Value).Ticks);
        }
    }
}
