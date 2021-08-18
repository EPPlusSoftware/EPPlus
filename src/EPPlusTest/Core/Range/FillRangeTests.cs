using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Globalization;
using System.Threading;

namespace EPPlusTest.Core.Range.Fill
{
    [TestClass]
    public class RangeFillTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _wsNum;
        static ExcelWorksheet _wsDate;
        static ExcelWorksheet _wsList;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Range_Fill.xlsx", true);
            _wsNum = _pck.Workbook.Worksheets.Add("FillNumbers");
            _wsDate = _pck.Workbook.Worksheets.Add("FillDates");
            _wsList = _pck.Workbook.Worksheets.Add("FillList");
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }

        [TestMethod]
        public void FillNumbers_WithStartAndStep()
        {
            _wsNum.Cells["A1:B5"].FillNumber(3, 2);
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
            _wsNum.Cells["C1:D5"].FillNumber();
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
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            _wsNum.Cells["A10:E11"].FillNumber(x =>
            {
                x.Direction = eFillDirection.Row;
                x.StartValue = 3;
                x.StepValue = 2;
                x.NumberFormat = "#,##0.00";
            }); 

            //Assert
            Assert.AreEqual(3D, _wsNum.Cells["A10"].Value);
            Assert.AreEqual(5D, _wsNum.Cells["B10"].Value);
            Assert.AreEqual(7D, _wsNum.Cells["C10"].Value);
            Assert.AreEqual(9D, _wsNum.Cells["D10"].Value);
            Assert.AreEqual(11D, _wsNum.Cells["E10"].Value);

            Assert.AreEqual(3D, _wsNum.Cells["A11"].Value);
            Assert.AreEqual(11D, _wsNum.Cells["E11"].Value);
            Assert.AreEqual("11.00", _wsNum.Cells["E11"].Text);
            Thread.CurrentThread.CurrentCulture = ci;
        }
        [TestMethod]
        public void FillNumbers_RowWithStartAndStep_EndValue()
        {
            _wsNum.Cells["A10:E11"].FillNumber(x =>
            {
                x.Direction = eFillDirection.Row;
                x.StartValue = 3;
                x.StepValue = 2;
                x.EndValue = 9;
            });

            //Assert
            Assert.AreEqual(3D, _wsNum.Cells["A10"].Value);
            Assert.AreEqual(5D, _wsNum.Cells["B10"].Value);
            Assert.AreEqual(7D, _wsNum.Cells["C10"].Value);
            Assert.AreEqual(9D, _wsNum.Cells["D10"].Value);
            Assert.IsNull(_wsNum.Cells["E10"].Value);

            Assert.AreEqual(3D, _wsNum.Cells["A11"].Value);
            Assert.IsNull(_wsNum.Cells["E11"].Value);
        }

        [TestMethod]
        public void FillNumbers_RowWithStartAndStepMultipy()
        {
            _wsNum.Cells["A13:E14"].FillNumber(x =>
            {
                x.CalculationMethod = eCalculationMethod.Multiply;
                x.Direction = eFillDirection.Row;
                x.StartValue = 3;
                x.StepValue = 2;
            });
            //Assert
            Assert.AreEqual(3D, _wsNum.Cells["A13"].Value);
            Assert.AreEqual(6D, _wsNum.Cells["B13"].Value);
            Assert.AreEqual(12D, _wsNum.Cells["C13"].Value);
            Assert.AreEqual(24D, _wsNum.Cells["D13"].Value);
            Assert.AreEqual(48D, _wsNum.Cells["E13"].Value);

            Assert.AreEqual(3D, _wsNum.Cells["A14"].Value);
            Assert.AreEqual(48D, _wsNum.Cells["E14"].Value);
        }
        [TestMethod]
        public void FillDate_WithStartAndStepDay()
        {
            var startDate = new DateTime(2021, 1, 1);
            _wsNum.Cells["A1:B5"].FillDateTime(new DateTime(2021,1,1));
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
            _wsDate.Cells["C1:D5"].FillDateTime();
            //Assert
            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsDate.Cells["C1"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay, ((DateTime)_wsDate.Cells["C2"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 2, ((DateTime)_wsDate.Cells["C3"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 3, ((DateTime)_wsDate.Cells["C4"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 4, ((DateTime)_wsDate.Cells["C5"].Value).Ticks);

            Assert.AreEqual("D1", _wsDate.Cells["D1"].Value);
            Assert.IsNull(_wsDate.Cells["D2"].Value);
        }
        [TestMethod]
        public void FillDate_StartDate()
        {
            var startDate = new DateTime(2021, 05, 29);
            _wsDate.Cells["E1"].Value = startDate;
            _wsDate.Cells["E1:F5"].FillDateTime(startDate);
            //Assert
            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsDate.Cells["E1"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay, ((DateTime)_wsDate.Cells["E2"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 2, ((DateTime)_wsDate.Cells["E3"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 3, ((DateTime)_wsDate.Cells["E4"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 4, ((DateTime)_wsDate.Cells["E5"].Value).Ticks);

            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsDate.Cells["F1"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 4, ((DateTime)_wsDate.Cells["F5"].Value).Ticks);
        }
        [TestMethod]
        public void FillDate_Week_WithNumberFormat()
        {
            var ci = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;

            var startDate = new DateTime(2021, 2, 15);
            _wsDate.Cells["G1"].Value = startDate;
            _wsDate.Cells["G1:H5"].FillDateTime(x=> { x.DateUnit = eDateUnit.Week;x.StartValue = startDate;x.NumberFormat = "yyyy-MM-dd"; });
            
            //Assert
            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsDate.Cells["G1"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 7, ((DateTime)_wsDate.Cells["G2"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 14, ((DateTime)_wsDate.Cells["G3"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 21, ((DateTime)_wsDate.Cells["G4"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 28, ((DateTime)_wsDate.Cells["G5"].Value).Ticks);

            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsDate.Cells["H1"].Value).Ticks);
            Assert.AreEqual(startDate.Ticks + TimeSpan.TicksPerDay * 28, ((DateTime)_wsDate.Cells["H5"].Value).Ticks);

            Assert.AreEqual("2021-02-22", _wsDate.Cells["G2"].Text);

            Thread.CurrentThread.CurrentCulture = ci;
        }
        [TestMethod]
        public void FillDate_Month_LastDayInMonth()
        {
            var startDate = new DateTime(2021, 2, 28);
            _wsDate.Cells["I1"].Value = startDate;
            _wsDate.Cells["I1:J5"].FillDateTime(x => { x.DateUnit = eDateUnit.Month; x.StartValue = startDate; });

            //Assert
            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsDate.Cells["I1"].Value).Ticks);
            Assert.AreEqual(new DateTime(2021, 3, 31).Ticks, ((DateTime)_wsDate.Cells["I2"].Value).Ticks);
            Assert.AreEqual(new DateTime(2021, 4, 30).Ticks, ((DateTime)_wsDate.Cells["I3"].Value).Ticks);
            Assert.AreEqual(new DateTime(2021, 5, 31).Ticks, ((DateTime)_wsDate.Cells["I4"].Value).Ticks);
            Assert.AreEqual(new DateTime(2021, 6, 30).Ticks, ((DateTime)_wsDate.Cells["I5"].Value).Ticks);

            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsDate.Cells["J1"].Value).Ticks);
            Assert.AreEqual(new DateTime(2021, 6, 30).Ticks, ((DateTime)_wsDate.Cells["J5"].Value).Ticks);
        }
        [TestMethod]
        public void FillDate_Month_LastDayInMonthNoWeekdays()
        {
            var startDate = new DateTime(2021, 1, 31);
            _wsDate.Cells["K1"].Value = startDate;
            _wsDate.Cells["K1:L5"].FillDateTime(x => { x.DateUnit = eDateUnit.Month; x.StartValue = startDate; x.WeekdaysOnly = true; });

            //Assert
            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsDate.Cells["K1"].Value).Ticks);
            Assert.AreEqual(new DateTime(2021, 2, 26).Ticks, ((DateTime)_wsDate.Cells["K2"].Value).Ticks);
            Assert.AreEqual(new DateTime(2021, 3, 31).Ticks, ((DateTime)_wsDate.Cells["K3"].Value).Ticks);
            Assert.AreEqual(new DateTime(2021, 4, 30).Ticks, ((DateTime)_wsDate.Cells["K4"].Value).Ticks);
            Assert.AreEqual(new DateTime(2021, 5, 31).Ticks, ((DateTime)_wsDate.Cells["K5"].Value).Ticks);

            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsDate.Cells["L1"].Value).Ticks);
            Assert.AreEqual(new DateTime(2021, 5, 31).Ticks, ((DateTime)_wsDate.Cells["L5"].Value).Ticks);
        }
        [TestMethod]
        public void FillDate_Month_LastDayInMonth_NoWeekdays_WithCalender()
        {
            var startDate = new DateTime(2021, 12, 17);
            _wsDate.Cells["M1"].Value = startDate;
            _wsDate.Cells["M1:N5"].FillDateTime(x => 
            { 
                x.DateUnit = eDateUnit.Week; 
                x.StartValue = startDate; 
                x.WeekdaysOnly = true;
                x.HolidayCalendar.UnionWith(new DateTime[] { new DateTime(2021, 12, 24), new DateTime(2021, 12, 25), new DateTime(2021, 12, 26), new DateTime(2021, 12, 31), new DateTime(2022, 01, 01) });
            });

            //Assert
            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsDate.Cells["M1"].Value).Ticks);
            Assert.AreEqual(new DateTime(2021, 12, 23).Ticks, ((DateTime)_wsDate.Cells["M2"].Value).Ticks);
            Assert.AreEqual(new DateTime(2021, 12, 30).Ticks, ((DateTime)_wsDate.Cells["M3"].Value).Ticks);
            Assert.AreEqual(new DateTime(2022, 1, 7).Ticks, ((DateTime)_wsDate.Cells["M4"].Value).Ticks);
            Assert.AreEqual(new DateTime(2022, 1, 14).Ticks, ((DateTime)_wsDate.Cells["M5"].Value).Ticks);

            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsDate.Cells["N1"].Value).Ticks);
            Assert.AreEqual(new DateTime(2022, 1, 14).Ticks, ((DateTime)_wsDate.Cells["N5"].Value).Ticks);
        }
        [TestMethod]
        public void FillDate_Year()
        {
            var startDate = new DateTime(2021, 2, 15);
            _wsDate.Cells["O1"].Value = startDate;
            _wsDate.Cells["O1:P5"].FillDateTime(x => { x.DateUnit = eDateUnit.Year; x.StartValue = startDate;x.EndValue = new DateTime(2024, 6, 30); });

            //Assert
            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsDate.Cells["O1"].Value).Ticks);
            Assert.AreEqual(new DateTime(2022, 2, 15).Ticks, ((DateTime)_wsDate.Cells["O2"].Value).Ticks);
            Assert.AreEqual(new DateTime(2023, 2, 15).Ticks, ((DateTime)_wsDate.Cells["O3"].Value).Ticks);
            Assert.AreEqual(new DateTime(2024, 2, 15).Ticks, ((DateTime)_wsDate.Cells["O4"].Value).Ticks);
            //Assert.AreEqual(new DateTime(2025, 2, 15).Ticks, ((DateTime)_wsDate.Cells["O5"].Value).Ticks);
            Assert.IsNull(_wsDate.Cells["O5"].Value);
            Assert.AreEqual(startDate.Ticks, ((DateTime)_wsDate.Cells["P1"].Value).Ticks);
            Assert.IsNull(_wsDate.Cells["P5"].Value);
        }
        [TestMethod]
        public void FillTime_Hours()
        {
            var startTime = DateTime.Parse("12:00:00");
            _wsDate.Cells["A20"].Value = startTime;
            _wsDate.Cells["A20:B24"].FillDateTime(x => { x.DateUnit = eDateUnit.Hour; x.StartValue = startTime; });

            //Assert
            Assert.AreEqual(startTime.Ticks, ((DateTime)_wsDate.Cells["A20"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerHour, ((DateTime)_wsDate.Cells["A21"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerHour * 2, ((DateTime)_wsDate.Cells["A22"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerHour * 3, ((DateTime)_wsDate.Cells["A23"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerHour * 4, ((DateTime)_wsDate.Cells["A24"].Value).Ticks);

            Assert.AreEqual(startTime.Ticks, ((DateTime)_wsDate.Cells["B20"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerHour * 4, ((DateTime)_wsDate.Cells["B24"].Value).Ticks);
        }
        [TestMethod]
        public void FillTime_Minutes()
        {
            var startTime = DateTime.Parse("00:45:00");
            _wsDate.Cells["C20"].Value = startTime;
            _wsDate.Cells["C20:D24"].FillDateTime(x => { x.DateUnit = eDateUnit.Minute; x.StartValue = startTime; });

            //Assert
            Assert.AreEqual(startTime.Ticks, ((DateTime)_wsDate.Cells["C20"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerMinute, ((DateTime)_wsDate.Cells["C21"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerMinute * 2, ((DateTime)_wsDate.Cells["C22"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerMinute * 3, ((DateTime)_wsDate.Cells["C23"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerMinute * 4, ((DateTime)_wsDate.Cells["C24"].Value).Ticks);

            Assert.AreEqual(startTime.Ticks, ((DateTime)_wsDate.Cells["D20"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerMinute * 4, ((DateTime)_wsDate.Cells["D24"].Value).Ticks);
        }
        [TestMethod]
        public void FillTime_Seconds()
        {
            var startTime = DateTime.Parse("00:00:30");
            _wsDate.Cells["E20"].Value = startTime;
            _wsDate.Cells["E20:F24"].FillDateTime(x => { x.DateUnit = eDateUnit.Second; x.StartValue = startTime; });

            //Assert
            Assert.AreEqual(startTime.Ticks, ((DateTime)_wsDate.Cells["E20"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerSecond, ((DateTime)_wsDate.Cells["E21"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerSecond * 2, ((DateTime)_wsDate.Cells["E22"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerSecond * 3, ((DateTime)_wsDate.Cells["E23"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerSecond * 4, ((DateTime)_wsDate.Cells["E24"].Value).Ticks);

            Assert.AreEqual(startTime.Ticks, ((DateTime)_wsDate.Cells["F20"].Value).Ticks);
            Assert.AreEqual(startTime.Ticks + TimeSpan.TicksPerSecond * 4, ((DateTime)_wsDate.Cells["F24"].Value).Ticks);
        }
        [TestMethod]
        public void FillList_Default()
        {
            var list = new string[] { "Monday","Tuesday","Wednesday" };
            _wsList.Cells["A1:B5"].FillList(list);

            //Assert
            Assert.AreEqual(list[0], _wsList.GetValue(1, 1));
            Assert.AreEqual(list[1], _wsList.GetValue(2, 1));
            Assert.AreEqual(list[2], _wsList.GetValue(3, 1));
            Assert.AreEqual(list[0], _wsList.GetValue(4, 1));
            Assert.AreEqual(list[1], _wsList.GetValue(5, 1));

            Assert.AreEqual(list[0], _wsList.GetValue(1, 2));
            Assert.AreEqual(list[1], _wsList.GetValue(2, 2));
            Assert.AreEqual(list[2], _wsList.GetValue(3, 2));
            Assert.AreEqual(list[0], _wsList.GetValue(4, 2));
            Assert.AreEqual(list[1], _wsList.GetValue(5, 2));
        }
    }
}
