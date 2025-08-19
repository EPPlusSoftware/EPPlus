using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.ArrayTests
{
    [TestClass]
    public class DateTimeFunctionsArrayTests
    {
        [TestMethod]
        public void TimeValueShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = "02:10";
                sheet.Cells["A2"].Value = "02:10 pm";
                sheet.Cells["A3"].Value = "03:15 am";
                sheet.Cells["B1:B3"].CreateArrayFormula("TIMEVALUE(A1:A3,3)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                var v1 = System.Math.Round((double)sheet.Cells["B1"].Value, 3);
                var v2 = System.Math.Round((double)sheet.Cells["B2"].Value, 3);
                var v3 = System.Math.Round((double)sheet.Cells["B3"].Value, 3);
                Assert.AreEqual(0.09, v1);
                Assert.AreEqual(0.59, v2);
                Assert.AreEqual(0.135, v3);
            }
        }

        [TestMethod]
        public void YearShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = new DateTime(2023, 3, 10);
                sheet.Cells["A2"].Value = new DateTime(2022, 3, 10);
                sheet.Cells["A3"].Value = new DateTime(2021, 3, 10);
                sheet.Cells["B1:B3"].CreateArrayFormula("YEAR(A1:A3)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                Assert.AreEqual(2023, sheet.Cells["B1"].Value);
                Assert.AreEqual(2022, sheet.Cells["B2"].Value);
                Assert.AreEqual(2021, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void MonthShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = new DateTime(2023, 3, 10);
                sheet.Cells["A2"].Value = new DateTime(2022, 4, 10);
                sheet.Cells["A3"].Value = new DateTime(2021, 5, 10);
                sheet.Cells["B1:B3"].CreateArrayFormula("MONTH(A1:A3)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                Assert.AreEqual(3, sheet.Cells["B1"].Value);
                Assert.AreEqual(4, sheet.Cells["B2"].Value);
                Assert.AreEqual(5, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void DayShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = new DateTime(2023, 3, 10);
                sheet.Cells["A2"].Value = new DateTime(2022, 4, 11);
                sheet.Cells["A3"].Value = new DateTime(2021, 5, 12);
                sheet.Cells["B1:B3"].CreateArrayFormula("DAY(A1:A3)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                Assert.AreEqual(10, sheet.Cells["B1"].Value);
                Assert.AreEqual(11, sheet.Cells["B2"].Value);
                Assert.AreEqual(12, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void HourShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = new DateTime(2023, 3, 10, 11, 30, 0);
                sheet.Cells["A2"].Value = new DateTime(2022, 4, 11, 12, 30, 0);
                sheet.Cells["A3"].Value = new DateTime(2021, 5, 12, 13, 30, 0);
                sheet.Cells["B1:B3"].CreateArrayFormula("HOUR(A1:A3)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                Assert.AreEqual(11, sheet.Cells["B1"].Value);
                Assert.AreEqual(12, sheet.Cells["B2"].Value);
                Assert.AreEqual(13, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void MinuteShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = new DateTime(2023, 3, 10, 11, 31, 0);
                sheet.Cells["A2"].Value = new DateTime(2022, 4, 11, 12, 32, 0);
                sheet.Cells["A3"].Value = new DateTime(2021, 5, 12, 13, 33, 0);
                sheet.Cells["B1:B3"].CreateArrayFormula("MINUTE(A1:A3)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                Assert.AreEqual(31, sheet.Cells["B1"].Value);
                Assert.AreEqual(32, sheet.Cells["B2"].Value);
                Assert.AreEqual(33, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void SecondShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");

                sheet.Cells["A1"].Value = new DateTime(2023, 3, 10, 11, 31, 31);
                sheet.Cells["A2"].Value = new DateTime(2022, 4, 11, 12, 32, 32);
                sheet.Cells["A3"].Value = new DateTime(2021, 5, 12, 13, 33, 33);
                sheet.Cells["B1:B3"].CreateArrayFormula("SECOND(A1:A3)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                Assert.AreEqual(31, sheet.Cells["B1"].Value);
                Assert.AreEqual(32, sheet.Cells["B2"].Value);
                Assert.AreEqual(33, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void WeekdayShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = new DateTime(2023, 3, 9, 1, 1, 1).ToOADate();
                sheet.Cells["A2"].Value = new DateTime(2023, 3, 10, 1, 1, 1).ToOADate();
                sheet.Cells["A3"].Value = new DateTime(2023, 3, 11, 1, 1, 1).ToOADate();
                sheet.Cells["B1:B3"].CreateArrayFormula("WEEKDAY(A1:A3,11)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                Assert.AreEqual(4, sheet.Cells["B1"].Value);
                Assert.AreEqual(5, sheet.Cells["B2"].Value);
                Assert.AreEqual(6, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void IsoWeeknumShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = new DateTime(2023, 3, 9, 1, 1, 1).ToOADate();
                sheet.Cells["A2"].Value = new DateTime(2023, 4, 10, 1, 1, 1).ToOADate();
                sheet.Cells["A3"].Value = new DateTime(2023, 7, 11, 1, 1, 1).ToOADate();
                sheet.Cells["B1:B3"].CreateArrayFormula("ISOWEEKNUM(A1:A3,11)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                Assert.AreEqual(10, sheet.Cells["B1"].Value);
                Assert.AreEqual(15, sheet.Cells["B2"].Value);
                Assert.AreEqual(28, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void DateValueShouldReturnVerticalArray()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = new DateTime(2023, 3, 15).ToString();
                sheet.Cells["A2"].Value = new DateTime(2023, 4, 15).ToString();
                sheet.Cells["A3"].Value = new DateTime(2023, 5, 15).ToString();
                sheet.Cells["B1:B3"].CreateArrayFormula("DATEVALUE(A1:A3)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                Assert.AreEqual(45000d, sheet.Cells["B1"].Value);
                Assert.AreEqual(45031d, sheet.Cells["B2"].Value);
                Assert.AreEqual(45061d, sheet.Cells["B3"].Value);
            }
        }

        [TestMethod]
        public void DaysShouldReturnVerticalArray_1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = new DateTime(2023, 3, 15).ToString();
                sheet.Cells["A2"].Value = new DateTime(2023, 4, 15).ToString();
                sheet.Cells["A3"].Value = new DateTime(2023, 5, 15).ToString();
                sheet.Cells["B1"].Value = new DateTime(2023, 3, 16).ToString();
                sheet.Cells["B2"].Value = new DateTime(2023, 4, 17).ToString();
                sheet.Cells["B3"].Value = new DateTime(2023, 5, 18).ToString();
                sheet.Cells["C1:C3"].CreateArrayFormula("DAYS(A1:A3,B1:B3)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                Assert.AreEqual(-1d, sheet.Cells["C1"].Value);
                Assert.AreEqual(-2d, sheet.Cells["C2"].Value);
                Assert.AreEqual(-3d, sheet.Cells["C3"].Value);
            }
        }

        [TestMethod]
        public void Days360ShouldReturnVerticalArray_1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = new DateTime(2023, 3, 9).ToString();
                sheet.Cells["A2"].Value = new DateTime(2023, 4, 10).ToString();
                sheet.Cells["A3"].Value = new DateTime(2023, 7, 11).ToString();
                sheet.Cells["B1"].Value = new DateTime(2023, 3, 10).ToString();
                sheet.Cells["B2"].Value = new DateTime(2023, 4, 15).ToString();
                sheet.Cells["B3"].Value = new DateTime(2023, 8, 12).ToString();
                sheet.Cells["C1:C3"].CreateArrayFormula("DAYS360(A1:A3,B1:B3)");
                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                Assert.AreEqual(1, sheet.Cells["C1"].Value);
                Assert.AreEqual(5, sheet.Cells["C2"].Value);
                Assert.AreEqual(31, sheet.Cells["C3"].Value);
            }
        }

        [TestMethod]
        public void DateDifShouldReturnVerticalArray_1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet1");
                sheet.Cells["A1"].Value = new DateTime(2023, 3, 9).ToString();
                sheet.Cells["A2"].Value = new DateTime(2023, 4, 10).ToString();
                sheet.Cells["A3"].Value = new DateTime(2023, 7, 11).ToString();
                sheet.Cells["B1"].Value = new DateTime(2023, 3, 10).ToString();
                sheet.Cells["B2"].Value = new DateTime(2023, 4, 15).ToString();
                sheet.Cells["B3"].Value = new DateTime(2023, 8, 12).ToString();
                sheet.Cells["C1"].Value = "d";
                sheet.Cells["C2"].Value = "d";
                sheet.Cells["C3"].Value = "d";
                sheet.Cells["D1:D3"].CreateArrayFormula("DATEDIF(A1:A3,B1:B3,C1:C3)");

                sheet.Calculate(x => x.PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel);
                Assert.AreEqual(1d, sheet.Cells["D1"].Value);
                Assert.AreEqual(5d, sheet.Cells["D2"].Value);
                Assert.AreEqual(32d, sheet.Cells["D3"].Value);
            }
        }
    }
}
