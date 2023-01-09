using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class IndexTests
    {
        private const string InputWorksheet = "Inputs";
        private const string ExtractWorksheet = "Test Formulas";
        /*
        [TestCase(1.0, 1.0, 100.1)]
        [TestCase(1.0, 2.0, 200.1)]
        [TestCase(12.1, 1.0, 101.20)]
        [TestCase(12.4, 1.0, 101.20)]
        [TestCase(12.5, 1.0, 101.20)]
        [TestCase(12.6, 1.0, 101.20)]
        [TestCase(12.9, 1.0, 101.20)]
       

        [TestCase(13, 1.0, 101.30)]
        [TestCase(13, 1.9, 101.30)]
        [TestCase(12.1, 2.0, 201.20)]
        [TestCase(12.4, 2.0, 201.20)]
        [TestCase(12.5, 2.0, 201.20)]
        [TestCase(12.6, 2.0, 201.20)]
        [TestCase(12.9, 2.0, 201.20)]
        [TestCase(13, 2.0, 201.30)]
         */

        [TestMethod]
        public void Test1()
        {
            Index(1.0, 1.0, 100.1);
        }

        [TestMethod]
        public void Test2()
        {
            Index(1.0, 2.0, 200.1);
        }

        [TestMethod]
        public void Test3()
        {
            Index(12.1, 1.0, 101.20);
        }

        [TestMethod]
        public void Test4()
        {
            Index(12.4, 1.0, 101.20);
        }

        [TestMethod]
        public void Test5()
        {
            Index(12.5, 1.0, 101.20);
        }

        [TestMethod]
        public void Test6()
        {
            Index(12.6, 1.0, 101.20);
        }

        [TestMethod]
        public void Test7()
        {
            Index(12.9, 1.0, 101.20);
        }

        [TestMethod]
        public void Test8()
        {
            Index(13, 1.0, 101.30);
        }

        [TestMethod]
        public void Test9()
        {
            Index(13, 1.0, 101.30);
        }

        [TestMethod]
        public void Test10()
        {
            Index(13, 1.9, 101.30);
        }

        [TestMethod]
        public void Test11()
        {
            Index(12.6, 2.0, 201.20);
        }

        [TestMethod]
        public void Test12()
        {
            Index(13, 2.0, 201.30);
        }

        public void Index(double row, double column, double expectedValue)
        {
            using (var package = CreateExcelPackage())
            {
                package.Workbook.Worksheets[InputWorksheet].Cells["B1"].Value = row;
                package.Workbook.Worksheets[InputWorksheet].Cells["B2"].Value = column;
                package.Workbook.Calculate(
                     new ExcelCalculationOption
                     {
                         PrecisionAndRoundingStrategy = PrecisionAndRoundingStrategy.Excel
                     });
                var extractedIndexValue = package.Workbook.Worksheets[ExtractWorksheet].Cells["B1"].Value;
                Assert.AreEqual(expectedValue, System.Math.Round((double)extractedIndexValue, 10));
            }

        }

        private static ExcelPackage CreateExcelPackage()
        {
            var package = new ExcelPackage() { Compression = CompressionLevel.Level0 };

            var sheet1 = package.Workbook.Worksheets.Add("LookupData");
            var b = 100d;
            var c = 200d;
            for(var x = 1; x < 16; x++)
            {
                b += 0.1;
                c += 0.1;
                sheet1.Cells["A" + x].Value = x;
                sheet1.Cells["B" + x].Value = b;
                sheet1.Cells["C" + x].Value = c;
            }
            var sheet2 = package.Workbook.Worksheets.Add(ExtractWorksheet);
            sheet2.Cells["B1:C15"].Formula = "INDEX(Sample_Index_Data,Row_Pos,Col_Pos)";

            var sheet3 = package.Workbook.Worksheets.Add(InputWorksheet);

            package.Workbook.Names.Add("Col_Pos", sheet3.Cells["B2"]);
            package.Workbook.Names.Add("Row_Pos", sheet3.Cells["B1"]);
            package.Workbook.Names.Add("Sample_Index_Data", sheet1.Cells["B1:C15"]);
            return package;
        }

    } 
}
