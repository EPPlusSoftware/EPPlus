using System.Globalization;
using System.Linq;
using System.Threading;

namespace EPPlusTest
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using OfficeOpenXml;
    using System;
    using System.IO;


    [TestClass]
    public class AmanaIssues : TestBase
    {
        [TestMethod]
        public void ExcelPackage_SaveAs_doesnt_throw_exception()
        {
            // Arrange
            var input = GetTestStream("SN_T_1506944663_AufrissGewinnundVerlustrechnung.xlsx");
            var package = new ExcelPackage(input);
            var output = Path.GetTempFileName();

            // Act-Assert
            package.SaveAs(output);

            // Cleanup
            File.Delete(output);

        }

        [TestMethod]
        public void Test_issue_with_whitespace_in_chart_xml()
        { 
            //Arrange
#if Core
            var dir = AppContext.BaseDirectory;
            dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
            var excelPackage = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "TestDoc_CountBlankSingleCell_xlsx.xlsx")));

            //Act
            var savePath = Path.Combine(TestContext.TestDeploymentDir, $"{TestContext.TestName}.xlsx");
            excelPackage.SaveAs(new FileInfo(savePath));

            excelPackage.Workbook.Calculate();

            //Asserts
            Assert.AreEqual("b)", excelPackage.Workbook.Worksheets[0].Cells["B3"].Value);
        }

        [TestMethod,
         Description(
             "Issue: If a cell is rich text and gets referenced by another cell by formula the Cell gets the Xml-Node as Value")]
        public void IssueTableWithXmlTags()
        {
            //Arrange
#if Core
            var dir = AppContext.BaseDirectory;
            dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
            //Act & Asserts
            var excelPackage = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "TestDoc_XMLTagsTable_xlsx.xlsx")));

            var sheet = excelPackage.Workbook.Worksheets["Tabelle1"];
            Assert.AreEqual(sheet.Cells["A1"].Value, sheet.Cells["B1"].Value);

            sheet.Calculate();
            Assert.AreEqual(sheet.Cells["A1"].Value, sheet.Cells["B1"].Value);

        }

        [TestMethod,
         Description(
             "Issue: If a VLookUp-Function contains a Date-Funktion as searchedValue an InvalidCastException is Thrown resulting in an #Value-Result")]
        public void TestIssueWithVLookUpDateValue()
        {
            //Arrange
#if Core
                var dir = AppContext.BaseDirectory;
                dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
            var excelPackage =
                new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "TestDoc_VLookUpDateValue_xlsx.xlsx")));

            //Act
            var worksheet = excelPackage.Workbook.Worksheets[0];

            worksheet.Calculate();

            //Assert
            Assert.AreEqual(worksheet.Cells["C2"].Value, worksheet.Cells["E3"].Value);
        }

        [TestMethod]
        public void Calculate_calculates_formula_with_external_link()
        {
            // Arrange
            var input = GetTestStream("ExternalReferences.xlsx");
            var package = new ExcelPackage(input);
            var sheet = package.Workbook.Worksheets[0];

            // Act
            sheet.Calculate();

            // Assert
            Assert.AreEqual(60d, sheet.Cells["A1"].Value);
            Assert.AreEqual(60d, sheet.Cells["A2"].Value);
            Assert.AreEqual(23d, sheet.Cells["B19"].Value);
            Assert.AreEqual(23d, sheet.Cells["B20"].Value);
        }

        [TestMethod]
        [DataRow("A1:A3,A5,A6,A7,A8,A10,A9,A11", ";A1;A2;A3;A5;A6;A7;A8;A10;A9;A11", 10)]
        [DataRow("A1", ";A1", 1)]
        [DataRow("A1:A4,A5:A7,A8:A11", ";A1;A2;A3;A4;A5;A6;A7;A8;A9;A10;A11", 11)]
        [DataRow("A1:A4,A5,A6,A7", ";A1;A2;A3;A4;A5;A6;A7", 7)]
        [DataRow("A1,A2,A3,A4:A7", ";A1;A2;A3;A4;A5;A6;A7", 7)]
        [DataRow("A1:A7", ";A1;A2;A3;A4;A5;A6;A7", 7)]
        [DataRow("A1,A2,A3,A4", ";A1;A2;A3;A4", 4)]
        [DataRow("A1,A2,A3:A5,A6,A7", ";A1;A2;A3;A4;A5;A6;A7", 7)]
        public void Cell_Range(string cellRange, string expectedAddresses, int expectedCount)
        {
            // Arrange
            var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("first");
            var sheet = package.Workbook.Worksheets.First();

            for (var i = 1; i <= 12; i++)
            {
                sheet.Cells[$"A{i}"].Value = 1;
            }

            sheet.Cells["A12"].Formula = "SUM(A1:A3,A5,A6,A7,A8,A10,A9,A11)";
            var counterFirstIteration = 0;
            var cellsFirstIteration = string.Empty;

            // Act
            var range = sheet.Cells[cellRange];
            foreach (var cell in range)
            {
                counterFirstIteration++;
                cellsFirstIteration = $"{cellsFirstIteration};{cell.Address}";
            }

            // Assert
            Assert.AreEqual(expectedAddresses, cellsFirstIteration);
            Assert.AreEqual(expectedCount, counterFirstIteration);
        }
        
        [TestMethod]
        public void IssueGermanBuildInNumberFormat()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");

            //Issue: The German BuildInNumberFormat differs from the English BuildInNumberformat therefore Epplus has to check the culture before parsing the id to NumberFormatExpression.
            // var excelTestFile = Resources.GermanBuildInNumberFormat;
            // excelStream.Write(excelTestFile, 0, excelTestFile.Length);

            var exlPackage = new ExcelPackage(GetTestStream("GermanBuildInNumberFormat.xlsx"));

            var ws = exlPackage.Workbook.Worksheets[0];

            var excelFormatString_2 = ws.Cells[2, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("General", excelFormatString_2);

            var excelFormatString_3 = ws.Cells[3, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("0", excelFormatString_3);

            var excelFormatString_4 = ws.Cells[4, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("0.00", excelFormatString_4);

            var excelFormatString_5 = ws.Cells[5, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("#,##0", excelFormatString_5);

            var excelFormatString_6 = ws.Cells[6, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("#,##0.00", excelFormatString_6);

            var excelFormatString_7 = ws.Cells[7, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("#,##0 _€;-#,##0 _€", excelFormatString_7);

            var excelFormatString_8 = ws.Cells[8, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("#,##0 _€;[Red]-#,##0 _€", excelFormatString_8);

            var excelFormatString_9 = ws.Cells[9, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("#,##0.00 _€;-#,##0.00 _€", excelFormatString_9);

            var excelFormatString_10 = ws.Cells[10, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("#,##0.00 _€;[Red]-#,##0.00 _€", excelFormatString_10);

            var excelFormatString_11 = ws.Cells[11, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("#,##0\\ \"€\";\\-#,##0\\ \"€\"", excelFormatString_11);

            var excelFormatString_12 = ws.Cells[12, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("#,##0\\ \"€\";[Red]\\-#,##0\\ \"€\"", excelFormatString_12);

            var excelFormatString_13 = ws.Cells[13, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("#,##0.00\\ \"€\";\\-#,##0.00\\ \"€\"", excelFormatString_13);

            var excelFormatString_14 = ws.Cells[14, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("#,##0.00\\ \"€\";[Red]\\-#,##0.00\\ \"€\"", excelFormatString_14);

            var excelFormatString_15 = ws.Cells[15, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("0%", excelFormatString_15);

            var excelFormatString_16 = ws.Cells[16, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("0.00%", excelFormatString_16);

            var excelFormatString_17 = ws.Cells[17, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("0.00E+00", excelFormatString_17);

            var excelFormatString_18 = ws.Cells[18, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("##0.0E+0", excelFormatString_18);

            var excelFormatString_19 = ws.Cells[19, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("# ?/?", excelFormatString_19);

            var excelFormatString_20 = ws.Cells[20, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("# ??/??", excelFormatString_20);

            var excelFormatString_21 = ws.Cells[21, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("dd.mm.yyyy", excelFormatString_21);

            var excelFormatString_22 = ws.Cells[22, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("dd. mm yy", excelFormatString_22);

            var excelFormatString_23 = ws.Cells[23, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("dd. mmm", excelFormatString_23);

            var excelFormatString_24 = ws.Cells[24, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("mmm yy", excelFormatString_24);

            var excelFormatString_25 = ws.Cells[25, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("h:mm AM/PM", excelFormatString_25);

            var excelFormatString_26 = ws.Cells[26, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("h:mm:ss AM/PM", excelFormatString_26);

            var excelFormatString_27 = ws.Cells[27, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("hh:mm", excelFormatString_27);

            var excelFormatString_28 = ws.Cells[28, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("hh:mm:ss", excelFormatString_28);

            var excelFormatString_29 = ws.Cells[29, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("dd.mm.yyyy hh:mm", excelFormatString_29);

            var excelFormatString_30 = ws.Cells[30, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("mm:ss", excelFormatString_30);

            var excelFormatString_31 = ws.Cells[31, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("mm:ss.0", excelFormatString_31);

            var excelFormatString_32 = ws.Cells[32, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("@", excelFormatString_32);

            var excelFormatString_33 = ws.Cells[33, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("[h]:mm:ss", excelFormatString_33);

            var excelFormatString_34 = ws.Cells[34, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("_-* #,##0\\ \"€\"_-;\\-* #,##0\\ \"€\"_-;_-* \"-\"\\ \"€\"_-;_-@_-",
                excelFormatString_34);

            var excelFormatString_35 = ws.Cells[35, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("_-* #,##0\\ _€_-;\\-* #,##0\\ _€_-;_-* \"-\"\\ _€_-;_-@_-",
                excelFormatString_35);

            var excelFormatString_36 = ws.Cells[36, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("_-* #,##0.00\\ \"€\"_-;\\-* #,##0.00\\ \"€\"_-;_-* \"-\"??\\ \"€\"_-;_-@_-",
                excelFormatString_36);

            var excelFormatString_37 = ws.Cells[37, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("_-* #,##0.00\\ _€_-;\\-* #,##0.00\\ _€_-;_-* \"-\"??\\ _€_-;_-@_-",
                excelFormatString_37);

            var excelFormatString_38 = ws.Cells[38, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("mmm\\ yyyy", excelFormatString_38);

            var excelFormatString_39 = ws.Cells[39, 1].Style?.Numberformat?.Format;
            Assert.AreEqual("[$-407]dddd\\,\\ d/\\ mmmm\\ yyyy", excelFormatString_39);
        }
        
        [TestMethod]
        public void Format()
        {
            var book = new ExcelPackage(new FileInfo(@"C:\temp\format.xlsx"));
            var sheet = book.Workbook.Worksheets[0];
            
            var format = sheet.Cells[1, 1].Style.Numberformat.Format; // format = "#,##0.00;(#,##0.00)"
        }
    }
}