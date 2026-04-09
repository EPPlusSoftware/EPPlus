using FakeItEasy.Configuration;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering.Conversions;


namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class GroupByTests : TestBase
    {

        [TestMethod]
        public void GroupBy()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Joe";
                s.Cells["A2"].Value = "Anna";
                s.Cells["A3"].Value = "Bertil";
                s.Cells["A4"].Value = "Joe";
                s.Cells["B1"].Value = 1;
                s.Cells["B2"].Value = 2;
                s.Cells["B3"].Value = 3;
                s.Cells["B4"].Value = 0;
                s.Cells["C1"].Formula = "GROUPBY(A1:A4, B1:B4, _xleta.SUM)";
                s.Calculate();
                Assert.AreEqual("Anna", s.Cells["C1"].Value);
                Assert.AreEqual("Bertil", s.Cells["C2"].Value);
                Assert.AreEqual("Joe", s.Cells["C3"].Value);
                Assert.AreEqual(2d, s.Cells["D1"].Value);
                Assert.AreEqual(3d, s.Cells["D2"].Value);
                Assert.AreEqual(1d, s.Cells["D3"].Value);
                Assert.AreEqual("Total", s.Cells["C4"].Value);
                Assert.AreEqual(6d, s.Cells["D4"].Value);
            }
        }

        [TestMethod]
        public void GroupByLambda()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Joe";
                s.Cells["A2"].Value = "Anna";
                s.Cells["A3"].Value = "Bertil";
                s.Cells["A4"].Value = "Joe";
                s.Cells["B1"].Value = 1;
                s.Cells["B2"].Value = 2;
                s.Cells["B3"].Value = 3;
                s.Cells["B4"].Value = 0;
                s.Cells["C1"].Formula = "GROUPBY(A1:A4, B1:B4, LAMBDA(x,SUM(x *2/3)) )";
                s.Calculate();
                Assert.AreEqual("Anna", s.Cells["C1"].Value);
                Assert.AreEqual("Bertil", s.Cells["C2"].Value);
                Assert.AreEqual("Joe", s.Cells["C3"].Value);                
                Assert.AreEqual(2d, s.Cells["D2"].Value);
                Assert.AreEqual(0.6667d, System.Math.Round((double)s.Cells["D3"].Value, 4));
                Assert.AreEqual("Total", s.Cells["C4"].Value);
                Assert.AreEqual(4d, s.Cells["D4"].Value);
            }
        }

        [TestMethod]
        public void GroupByFieldHeaders()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "A";
                s.Cells["A2"].Value = "B"; 
                s.Cells["A3"].Value = "A";
                s.Cells["A4"].Value = "A";
                s.Cells["A5"].Value = "B";

                s.Cells["B1"].Value = "X";
                s.Cells["B2"].Value = "X";
                s.Cells["B3"].Value = "Y";
                s.Cells["B4"].Value = "Y";
                s.Cells["B5"].Value = "Y";

                s.Cells["C1"].Value = 1;
                s.Cells["C2"].Value = 1;
                s.Cells["C3"].Value = 2;
                s.Cells["C4"].Value = 1;
                s.Cells["C5"].Value = 1;

                s.Cells["D1"].Formula = "GROUPBY(A1:B5, C1:C5, _xleta.SUM,,2)";
                s.Calculate();

                Assert.AreEqual("A", s.Cells["D1"].Value);
                Assert.AreEqual("X", s.Cells["E1"].Value);
                Assert.AreEqual(1d, s.Cells["F1"].Value);

                Assert.AreEqual("A", s.Cells["D2"].Value);
                Assert.AreEqual("Y", s.Cells["E2"].Value);
                Assert.AreEqual(3d, s.Cells["F2"].Value);
                // Subtotal row
                Assert.AreEqual("A", s.Cells["D3"].Value);
                Assert.AreEqual(4d, s.Cells["F3"].Value);

                Assert.AreEqual("B", s.Cells["D4"].Value);
                Assert.AreEqual("X", s.Cells["E4"].Value);
                Assert.AreEqual(1d, s.Cells["F4"].Value);
                Assert.AreEqual("B", s.Cells["D5"].Value);
                Assert.AreEqual("Y", s.Cells["E5"].Value);
                Assert.AreEqual(1d, s.Cells["F5"].Value);

                Assert.AreEqual("B", s.Cells["D6"].Value);
                Assert.AreEqual(2d, s.Cells["F6"].Value);

                Assert.AreEqual("Grand Total", s.Cells["D7"].Value);
                Assert.AreEqual(6d, s.Cells["F7"].Value);
            }
        }

        [TestMethod]
        public void GroupByFilteredArray()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Anna";
                s.Cells["A2"].Value = "Joe";
                s.Cells["A3"].Value = "Bertil";
                s.Cells["A4"].Value = "ANNA";
                s.Cells["A5"].Value = "Anna";
                s.Cells["A6"].Value = "Bertil";
                s.Cells["A7"].Value = "Anna";
                s.Cells["A8"].Value = "Joe";

                s.Cells["B1"].Value = 1;
                s.Cells["B2"].Value = 1;
                s.Cells["B3"].Value = 1;

                s.Cells["B4"].Value = 2;
                s.Cells["B5"].Value = 2;
                s.Cells["B7"].Value = 3;

                s.Cells["C1"].Formula = "GROUPBY(A1:A8, B1:B8, _xleta.SUM,,,,A1:A8 =\"ANNA\")";
                s.Calculate();
                Assert.AreEqual(s.Cells["C1"].Value, "Anna");
                Assert.AreEqual(8d, s.Cells["D1"].Value);
            }
        }

        [TestMethod]
        public void GroupByFieldRelationship()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = new DateTime(2025, 01, 01);
                s.Cells["A2"].Value = new DateTime(2025, 02, 01);
                s.Cells["A3"].Value = new DateTime(2025, 03, 01);
                s.Cells["A4"].Value = new DateTime(2025, 03, 01);
                s.Cells["A5"].Value = new DateTime(2025, 01, 01);

                s.Cells["B1"].Value = 30;
                s.Cells["B2"].Value = 20;
                s.Cells["B3"].Value = 54;
                s.Cells["B4"].Value = 54;
                s.Cells["B5"].Value = 23;

                s.Cells["C1"].Formula = "=GROUPBY(HSTACK(CHOOSE(MONTH(A1:A5),\"Jan\",\"Feb\",\"Mar\",\"Apr\",\"Maj\",\"Jun\",\"Jul\",\"Aug\",\"Sep\",\"Okt\",\"Nov\",\"Dec\"), MONTH(A1:A5) ), B1:B5, _xleta.SUM,,,2,,1 )";
                s.Calculate();
                Assert.AreEqual("Jan", s.Cells["C1"].Value);
                Assert.AreEqual(53d, s.Cells["E1"].Value);
                Assert.AreEqual("Feb", s.Cells["C2"].Value);
                Assert.AreEqual(20d, s.Cells["E2"].Value);
                Assert.AreEqual("Mar", s.Cells["C3"].Value);
                Assert.AreEqual(108d, s.Cells["E3"].Value);
            }
        }

        [TestMethod]
        public void GroupByFieldRelationship2()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = new DateTime(2025, 01, 01);
                s.Cells["A2"].Value = new DateTime(2025, 02, 01);
                s.Cells["A3"].Value = new DateTime(2025, 03, 01);
                s.Cells["A4"].Value = new DateTime(2025, 03, 01);
                s.Cells["A5"].Value = new DateTime(2025, 01, 01);

                s.Cells["B1"].Value = 30;
                s.Cells["B2"].Value = 20;
                s.Cells["B3"].Value = 54;
                s.Cells["B4"].Value = 54;
                s.Cells["B5"].Value = 23;

                s.Cells["C1"].Formula = "=CHOOSECOLS(GROUPBY(HSTACK(CHOOSE(MONTH(A1:A5),\"Jan\",\"Feb\",\"Mar\",\"Apr\",\"Maj\",\"Jun\",\"Jul\",\"Aug\",\"Sep\",\"Okt\",\"Nov\",\"Dec\"), MONTH(A1:A5) ), B1:B5, _xleta.SUM,,,2,,1) ,{1,3})";
                s.Calculate();
                Assert.AreEqual("Jan", s.Cells["C1"].Value);
                Assert.AreEqual(53d, s.Cells["D1"].Value);
                Assert.AreEqual("Feb", s.Cells["C2"].Value);
                Assert.AreEqual(20d, s.Cells["D2"].Value);
                Assert.AreEqual("Mar", s.Cells["C3"].Value);
                Assert.AreEqual(108d, s.Cells["D3"].Value);
            }
        }
        [TestMethod]
        public void GroupBy_NoTotals_ShouldNotIncludeTotalRow()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Joe";
                s.Cells["A2"].Value = "Anna";
                s.Cells["A3"].Value = "Bertil";
                s.Cells["A4"].Value = "Joe";
                s.Cells["B1"].Value = 1;
                s.Cells["B2"].Value = 2;
                s.Cells["B3"].Value = 3;
                s.Cells["B4"].Value = 0;

                s.Cells["C1"].Formula = "GROUPBY(A1:A4, B1:B4, _xleta.SUM,, 0)";
                s.Calculate();

                Assert.AreEqual("Anna", s.Cells["C1"].Value);
                Assert.AreEqual("Bertil", s.Cells["C2"].Value);
                Assert.AreEqual("Joe", s.Cells["C3"].Value);
                Assert.AreEqual(2d, s.Cells["D1"].Value);
                Assert.AreEqual(3d, s.Cells["D2"].Value);
                Assert.AreEqual(1d, s.Cells["D3"].Value);

                Assert.AreNotEqual(s.Cells["C4"].Value, "Total");
                Assert.AreNotEqual(s.Cells["D4"].Value, 0d);
            }
        }

        [TestMethod]
        public void GroupBySortingMultipleCols()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Stockholm";
                s.Cells["A2"].Value = "Stockholm";
                s.Cells["A3"].Value = "Göteborg";
                s.Cells["A4"].Value = "Linköping";
                s.Cells["A5"].Value = "Linköping";
                s.Cells["A6"].Value = "Göteborg";
                s.Cells["A7"].Value = "Stockholm";

                s.Cells["B1"].Value = "Cykel";
                s.Cells["B2"].Value = "Boll";
                s.Cells["B3"].Value = "Fisk";
                s.Cells["B4"].Value = "Bomb";
                s.Cells["B5"].Value = "Bok";
                s.Cells["B6"].Value = "Boll";
                s.Cells["B7"].Value = "Fisk";

                s.Cells["C1"].Value = "Vällingby";
                s.Cells["C2"].Value = "Vällingby";
                s.Cells["C3"].Value = "Majorna";
                s.Cells["C4"].Value = "Skäggetorp";
                s.Cells["C5"].Value = "Tornby";
                s.Cells["C6"].Value = "Majorna";
                s.Cells["C7"].Value = "Vällingby";

                s.Cells["D1"].Value = 1000;
                s.Cells["D2"].Value = 300;
                s.Cells["D3"].Value = 200;
                s.Cells["D4"].Value = 3000;
                s.Cells["D5"].Value = 700;
                s.Cells["D6"].Value = 300;
                s.Cells["D7"].Value = 300;

                s.Cells["E1"].Formula = "GROUPBY(A1:C7, D1:D7, _xleta.SUM,0,2,-1,,0)";
                s.Calculate();

                Assert.AreEqual("Stockholm", s.Cells["E1"].Value);
                Assert.AreEqual("Grand Total", s.Cells["E11"].Value);
                Assert.AreEqual(5800d, s.Cells["H11"].Value);
            }
        }

        [TestMethod]
        public void GroupByTextFunction()
        {
            using (var package = new ExcelPackage())
            {
                SwitchToCulture();
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Kalle";
                s.Cells["A2"].Value = "Alice";
                s.Cells["A3"].Value = "Kalle";
                s.Cells["A4"].Value = "Alva";

                s.Cells["B1"].Value = "Hoppade";
                s.Cells["B2"].Value = "Sprang";
                s.Cells["B3"].Value = "Hoppade";
                s.Cells["B4"].Value = "Gick";

                s.Cells["C1"].Formula = "GROUPBY(A1:A4, B1:B4, _xleta.ARRAYTOTEXT)";
                s.Calculate();
                Assert.AreEqual("Hoppade; Sprang; Hoppade; Gick", s.Cells["D4"].Value);
                SwitchBackToCurrentCulture();
            }
        }

        [TestMethod]
        public void GroupByAVERAGE()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A1"].Value = "Kalle";
                s.Cells["A2"].Value = "Alice";
                s.Cells["A3"].Value = "Kalle";
                s.Cells["A4"].Value = "Alva";

                s.Cells["B1"].Value = 1;
                s.Cells["B2"].Value = 2;
                s.Cells["B3"].Value = 3;
                s.Cells["B4"].Value = 4;

                s.Cells["C1"].Formula = "GROUPBY(A1:A4, B1:B4, _xleta.AVERAGE)";
                s.Calculate();

                Assert.AreEqual("Total", s.Cells["C4"].Value);
                Assert.AreEqual(2.5d, s.Cells["D4"].Value);
            }
        }

        [TestMethod]
        public void GroupByShouldInsertZeroWhenEmptyAndNumericFunction()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");

                s.Cells["A1"].Value = "B";
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["A4"].Value = "A";
                s.Cells["A5"].Value = "C";

                s.Cells["B1"].Value = 1;
                s.Cells["B3"].Value = 3;
                s.Cells["B5"].Value = 4;

                s.Cells["C1"].Formula = "GROUPBY(A1:A5, B1:B5, _xleta.SUM)";
                s.Calculate();
                Assert.AreEqual(0d, s.Cells["D1"].Value);
            }
        }

        [TestMethod]
        public void GroupByMultipleFunctions()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");

                s.Cells["A1"].Value = "B";
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["A4"].Value = "A";
                s.Cells["A5"].Value = "C";

                s.Cells["B1"].Value = 1;
                s.Cells["B3"].Value = 3;
                s.Cells["B5"].Value = 4;

                s.Cells["C1"].Formula = "=GROUPBY(A1:A5, B1:B5,HSTACK(_xleta.COUNT, _xleta.SUM, _xleta.PERCENTOF))";
                s.Calculate();
                //Assert.AreEqual(null, s.Cells["C1"].Value);
                Assert.AreEqual("COUNT", s.Cells["D1"].Value);
                Assert.AreEqual("SUM", s.Cells["E1"].Value);
                Assert.AreEqual("PERCENTOF", s.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void GroupByMultipleFunctions2()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");

                s.Cells["A1"].Value = "Rubrik";
                s.Cells["A2"].Value = "B";
                s.Cells["A3"].Value = "A";
                s.Cells["A4"].Value = "B";
                s.Cells["A5"].Value = "A";
                s.Cells["A6"].Value = "C";

                s.Cells["B1"].Value = "Siffor";
                s.Cells["B2"].Value = 1;
                s.Cells["B4"].Value = 3;
                s.Cells["B6"].Value = 4;

                s.Cells["C1"].Formula = "=GROUPBY(A1:A6, B1:B6,HSTACK(_xleta.COUNT, _xleta.SUM, _xleta.PERCENTOF),3)";
                s.Calculate();
                //Assert.AreEqual(null, s.Cells["C1"].Value);
                Assert.AreEqual("COUNT", s.Cells["D1"].Value);
                Assert.AreEqual("SUM", s.Cells["E1"].Value);
                Assert.AreEqual("PERCENTOF", s.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void GroupByMultipleFunctions3()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");

                s.Cells["A1"].Value = "Rubrik";
                s.Cells["A2"].Value = "B";
                s.Cells["A3"].Value = "A";
                s.Cells["A4"].Value = "B";
                s.Cells["A5"].Value = "A";
                s.Cells["A6"].Value = "C";

                s.Cells["B1"].Value = "Siffor";
                s.Cells["B2"].Value = 1;
                s.Cells["B4"].Value = 3;
                s.Cells["B6"].Value = 4;    

                s.Cells["C1"].Formula = "=GROUPBY(A1:A6, B1:B6,VSTACK(_xleta.COUNT, _xleta.SUM, _xleta.PERCENTOF),3)";
                s.Calculate();

                Assert.AreEqual("COUNT", s.Cells["D2"].Value);
                Assert.AreEqual("SUM", s.Cells["D3"].Value);
                Assert.AreEqual("PERCENTOF", s.Cells["D4"].Value);
            }
        }        

        [TestMethod]
        public void GroupByMultipleFunctionsCustomLambda()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");

                s.Cells["A1"].Value = "Rubrik";
                s.Cells["A2"].Value = "B";
                s.Cells["A3"].Value = "A";
                s.Cells["A4"].Value = "B";
                s.Cells["A5"].Value = "A";
                s.Cells["A6"].Value = "C";

                s.Cells["B1"].Value = "Siffor";
                s.Cells["B2"].Value = 1;
                s.Cells["B4"].Value = 3;
                s.Cells["B6"].Value = 4;

                s.Cells["C1"].Formula = "GROUPBY(A1:A6, B1:B6,HSTACK(_xleta.COUNT, LAMBDA(x,SUM(x *2/3)), LAMBDA(x,SUM(x *2)) ),3)";
                //  LAMBDA(x, SUM(x*4/2)) LAMBDA(x,SUM(x *2/3))
                s.Calculate();

                Assert.AreEqual("COUNT", s.Cells["D1"].Value);
                Assert.AreEqual("CUSTOM1", s.Cells["E1"].Value);
                Assert.AreEqual("CUSTOM2", s.Cells["F1"].Value);               
            }
        }

        [TestMethod]
        public void GroupByMultipleFunctionsCustomLambda2()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");

                s.Cells["A1"].Value = "Rubrik";
                s.Cells["A2"].Value = "B";
                s.Cells["A3"].Value = "A";
                s.Cells["A4"].Value = "B";
                s.Cells["A5"].Value = "A";
                s.Cells["A6"].Value = "C";

                s.Cells["B1"].Value = "Siffor";
                s.Cells["B2"].Value = 1;
                s.Cells["B4"].Value = 3;
                s.Cells["B6"].Value = 4;

                s.Cells["C1"].Formula = "GROUPBY(A1:A6, B1:B6,HSTACK(_xleta.COUNT, LAMBDA(x,SUM(x *2/3)), _xleta.PERCENTOF, LAMBDA(x,SUM(x *2)) ) )";
                //  LAMBDA(x, SUM(x*4/2)) LAMBDA(x,SUM(x *2/3))
                s.Calculate();

                Assert.AreEqual("COUNT", s.Cells["D1"].Value);
                Assert.AreEqual("CUSTOM1", s.Cells["E1"].Value);
                Assert.AreEqual("PERCENTOF", s.Cells["F1"].Value);
                Assert.AreEqual("CUSTOM2", s.Cells["G1"].Value);
            }
        }

        [TestMethod]
        public void GroupBySortByArrayInput()
        {
            using (var package = new ExcelPackage())
            {
                var s = package.Workbook.Worksheets.Add("test");
                s.Cells["A2"].Value = "A";
                s.Cells["A3"].Value = "B";
                s.Cells["B2"].Value = "C";
                s.Cells["B3"].Value = "A";
                s.Cells["C2"].Value = "A";
                s.Cells["C3"].Value = "B";
                s.Cells["D2"].Value = "C";
                s.Cells["D3"].Value = "A";

                s.Cells["E2"].Value = 4;
                s.Cells["E3"].Value = 2;
                s.Cells["F2"].Value = 6;
                s.Cells["F3"].Value = 5;
                s.Cells["F6"].Formula = "GROUPBY(A2:D3,E2:F3,_xleta.SUM,,,{-1,2,3})";
                s.Calculate();

                Assert.AreEqual(s.Cells["F6"].Value, "B");
                Assert.AreEqual(s.Cells["J8"].Value, 6d);
                Assert.AreEqual(s.Cells["K8"].Value, 11d);
            }
        }

        // TESTA SKICKA IN LAMBDA SÅ ATT VI KAN SE ATT CUSTOM funktionerna FÅR RÄTT HEADERS "CUSTOM1, CUSTOM2..."
        // NOTE: Verkar vara nåt som blir knasigt med headers. Dem skrivs ut på fel ställe i rangen, ex: rubrik på fel ställe ovan
    }
}
