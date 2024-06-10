using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Globalization;
using System.IO;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class AmanaIssues : TestBase
    {
        [TestMethod]
        public void IssueMissingDecimalsTextFormular()
        {
            //Issue: TEXT-formular deletes decimals in german format
#if Core
            var dir = AppContext.BaseDirectory;
            dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
            using var exlPackage = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "Textformat.xlsx")));
            
            ExcelWorkbook.Culture = new CultureInfo("de-DE");
            exlPackage.Workbook.Calculate();

            Assert.AreEqual("292.336,30 €", exlPackage.Workbook.Worksheets[0].Cells["A1"].Text);
            Assert.AreEqual("292336,300000 €", exlPackage.Workbook.Worksheets[0].Cells["A2"].Value);
            Assert.AreEqual("292.336 €", exlPackage.Workbook.Worksheets[0].Cells["A3"].Value);
        
            Assert.AreEqual("292.336,30 €", exlPackage.Workbook.Worksheets[0].Cells["A5"].Value);
            Assert.AreEqual("-292336-- €", exlPackage.Workbook.Worksheets[0].Cells["A6"].Value);
            //Assert.AreEqual("-292336,--€", exlPackage.Workbook.Worksheets[0].Cells["A6"].Value);
            Assert.AreEqual("233.127,25 €)", exlPackage.Workbook.Worksheets[0].Cells["A7"].Value);
            Assert.AreEqual("-233127--€)", exlPackage.Workbook.Worksheets[0].Cells["A8"].Value);
            //Assert.AreEqual("-233127,--€)", exlPackage.Workbook.Worksheets[0].Cells["A8"].Value);
            Assert.AreEqual("0,00 €", exlPackage.Workbook.Worksheets[0].Cells["A9"].Value);
            Assert.AreEqual("--- €", exlPackage.Workbook.Worksheets[0].Cells["A10"].Value);
            //Assert.AreEqual("-,-- €", exlPackage.Workbook.Worksheets[0].Cells["A10"].Value);
            Assert.AreEqual("0,00 €)", exlPackage.Workbook.Worksheets[0].Cells["A11"].Value);
            Assert.AreEqual("---€)", exlPackage.Workbook.Worksheets[0].Cells["A12"].Value);
            //Assert.AreEqual("-,--€)", exlPackage.Workbook.Worksheets[0].Cells["A12"].Value);
            Assert.AreEqual("1.027,60 €", exlPackage.Workbook.Worksheets[0].Cells["A13"].Value);
            Assert.AreEqual("-1028-- €)", exlPackage.Workbook.Worksheets[0].Cells["A14"].Value);
            //Assert.AreEqual("-1028,--€)", exlPackage.Workbook.Worksheets[0].Cells["A14"].Value);
            Assert.AreEqual("445,58 €)", exlPackage.Workbook.Worksheets[0].Cells["A15"].Value);
            Assert.AreEqual("-446-- €)", exlPackage.Workbook.Worksheets[0].Cells["A16"].Value);
            //Assert.AreEqual("-446,--€)", exlPackage.Workbook.Worksheets[0].Cells["A16"].Value);
            Assert.AreEqual("0,00 €", exlPackage.Workbook.Worksheets[0].Cells["A17"].Value);
            Assert.AreEqual("0,00 €)", exlPackage.Workbook.Worksheets[0].Cells["A18"].Value);
            Assert.AreEqual("--- €)", exlPackage.Workbook.Worksheets[0].Cells["A19"].Value);
            //Assert.AreEqual("-,--€)", exlPackage.Workbook.Worksheets[0].Cells["A19"].Value);
            Assert.AreEqual("--- €", exlPackage.Workbook.Worksheets[0].Cells["A20"].Value);
            //Assert.AreEqual("-,--€", exlPackage.Workbook.Worksheets[0].Cells["A20"].Value);
        }
    }
}
