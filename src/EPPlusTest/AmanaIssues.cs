namespace EPPlusTest
{
    using EPPlusTest.Properties;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using OfficeOpenXml;
    using System;
    using System.IO;
    using System.Linq;

    [TestClass]
    public class AmanaIssues : TestBase
    {
        [TestMethod, 
         Description("If a cell contains a hyperlink with special characters such as ä,ö,ü Excel encodes the link not in UTF-8 to keep the rule that a target link must be shorter than 2080 characters")]
        public void Test_can_not_open_file_after_saving()
        {
            //Arrange
#if ! Core
         
            var dir = AppDomain.CurrentDomain.BaseDirectory;

            var excelPackage = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "TestDoc_CellWithHyperlink_xlsx.xlsx")));
            var ws = excelPackage.Workbook.Worksheets[0];

            var savePath = Path.Combine(TestContext.TestDeploymentDir, $"{TestContext.TestName}.xlsx");
            excelPackage.SaveAs(new FileInfo(savePath));
            var exApp = new Microsoft.Office.Interop.Excel.Application();

            try
            {
                var exWbk = exApp.Workbooks.Open(savePath);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Assert.Fail("It is not possible to open the workbook after EPPlus saved it.");
            }
            finally
            {
                exApp.Workbooks.Close();
            }
#endif
        }


        [TestMethod]
        public void Test_correct_values_in_WENNs_formula()
        {
            //Arrange
#if Core
            var dir = AppContext.BaseDirectory;
            dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
            var excelPackage = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "TestDoc_Wenns_Formula_xlsx.xlsx")));
            var ws = excelPackage.Workbook.Worksheets[0];

            //Act
            ws.Calculate();

            var value1 = ws.Cells["C3"].Value.ToString();
            var value2 = ws.Cells["C4"].Value.ToString();
            var value3 = ws.Cells["C5"].Value.ToString();
            var value4 = ws.Cells["C6"].Value.ToString();
            var value5 = ws.Cells["C7"].Value.ToString();

            //Asserts
            Assert.IsTrue(value1.Equals("one"));
            Assert.IsTrue(value2.Equals("two"));
            Assert.IsTrue(value3.Equals("three")); 
            Assert.IsTrue(value4.Equals("four"));
            Assert.IsTrue(value5.Equals("#N/A"));
        }


        [TestMethod, Description("If a formula contains external links the old value should be used instead of resulting in #NAME-Error")]
        public void Calculate_sets_old_value_if_formula_contains_external_link()
        {
            //Arrange
#if Core
            var dir = AppContext.BaseDirectory;
            dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
            var excelPackage = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "TestDoc_CellsWithFormulas_xlsx.xlsx")));
            var ws = excelPackage.Workbook.Worksheets[2];

            //Act
            ws.Calculate();

            //Asserts
            for (var i = 9; i <= 148; i++)
                Assert.AreEqual(ws.Cells[i, 3].Value, ws.Cells[i + 140, 3].Value);
        }


        [TestMethod, Description("If a formula contains external links the old value should be used instead of resulting in #NAME-Error")]
        public void Calculate_sets_old_value_if_formula_contains_external_link2()
        {

            //Arrange
#if Core
            var dir = AppContext.BaseDirectory;
            dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
            var excelPackage = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "TestDoc_WithExternalReferences_xlsx.xlsx")));

            //Act
            var ws = excelPackage.Workbook.Worksheets[0];
            ws.Calculate();

            //Asserts
            Assert.AreEqual(60d, ws.Cells["A1"].Value);
            Assert.AreEqual(60d, ws.Cells["A2"].Value);
            Assert.AreEqual(23d, ws.Cells["B19"].Value);
            Assert.AreEqual(23d, ws.Cells["B20"].Value);

        }

        [TestMethod]
        public void Test_roman_values()
        {
            //Arrange
#if Core
            var dir = AppContext.BaseDirectory;
            dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
            var excelPackage = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "TestDoc_WithRomanValues_xlsx.xlsx")));
            var ws = excelPackage.Workbook.Worksheets[0];

            //Act
            ws.Calculate();

            //Asserts
            //Parameter
            Assert.AreEqual(ws.Cells["A1"].Value, ws.Cells["B1"].Value);
            Assert.AreEqual(ws.Cells["A2"].Value, ws.Cells["B2"].Value);
            Assert.AreEqual(ws.Cells["A3"].Value, ws.Cells["B3"].Value);
            Assert.AreEqual(ws.Cells["A4"].Value, ws.Cells["B4"].Value);
            Assert.AreEqual(ws.Cells["A5"].Value, ws.Cells["B5"].Value);
            Assert.AreEqual(ws.Cells["A6"].Value, ws.Cells["B6"].Value);
            Assert.AreEqual(ws.Cells["A7"].Value, ws.Cells["B7"].Value);
            Assert.AreEqual(ws.Cells["A8"].Value, ws.Cells["B8"].Value);
            Assert.AreEqual(ws.Cells["A9"].Value, ws.Cells["B9"].Value);
            Assert.AreEqual(ws.Cells["A10"].Value, ws.Cells["B10"].Value);

            //Wrong Parameter
            Assert.AreEqual(ws.Cells["C1"].Value, ws.Cells["D1"].Value);
            Assert.AreEqual(ws.Cells["C2"].Value, ws.Cells["D2"].Value);
            Assert.AreEqual(ws.Cells["C3"].Value, ws.Cells["D3"].Value);
            Assert.AreEqual(ws.Cells["C4"].Value, ws.Cells["D4"].Value);
            Assert.AreEqual(ws.Cells["C5"].Value, string.Empty);
        }

        [TestMethod]
        public void Test_roman_values_for_excel_function()
        {
            //Arrange
#if Core
            var dir = AppContext.BaseDirectory;
            dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
            var excelPackage = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "TestDoc_WithRomanNumber_xlsx.xlsx")));
            var ws = excelPackage.Workbook.Worksheets[0];
            
            //Act
            ws.Calculate();

            //Asserts
            //no Parameter
            for (var i = 1; i <= ws.Cells["A:A"].Count(); i++)
                Assert.AreEqual(ws.Cells[i, 1].Value, ws.Cells[i, (1 + 11)].Value);

            //Parameter 0
            for (var i = 1; i <= ws.Cells["B:B"].Count(); i++)
                Assert.AreEqual(ws.Cells[i, 2].Value, ws.Cells[i, (2 + 11)].Value);
            //Parameter 1
            for (var i = 1; i <= ws.Cells["C:C"].Count(); i++)
                Assert.AreEqual(ws.Cells[i, 3].Value, ws.Cells[i, (3 + 11)].Value);
            //Parameter 2
            for (var i = 1; i <= ws.Cells["D:D"].Count(); i++)
                Assert.AreEqual(ws.Cells[i, 4].Value, ws.Cells[i, (4 + 11)].Value);
            //Parameter 3
            for (var i = 1; i <= ws.Cells["E:E"].Count(); i++)
                Assert.AreEqual(ws.Cells[i, 5].Value, ws.Cells[i, (5 + 11)].Value);
            //Parameter 4
            for (var i = 1; i <= ws.Cells["F:F"].Count(); i++)
                Assert.AreEqual(ws.Cells[i, 6].Value, ws.Cells[i, (6 + 11)].Value);
            //Parameter TRUE
            for (var i = 1; i <= ws.Cells["G:G"].Count(); i++)
                Assert.AreEqual(ws.Cells[i, 7].Value, ws.Cells[i, (7 + 11)].Value);
            //Parameter FALSE
            for (var i = 1; i <= ws.Cells["H:H"].Count(); i++)
                Assert.AreEqual(ws.Cells[i, 7].Value, ws.Cells[i, 7 + 11].Value);
        }



        [TestMethod,
         Description("If a Chart.xml contains ExtLst Nodes than the indentation of the chart.xml leads to corrupt Excel files")]
        public void IssueWhitespaceInChartXml()
        {
            /* Note: The Microsoft.Office.Interop.Excel library is not compatible with all .Core frameworks. */
            
            //Arrange
#if ! Core
            var dir = AppDomain.CurrentDomain.BaseDirectory;
            var excelPackage = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "TestDoc_WithCharts_xlsx.xlsx")));

            //Act
            var savePath = Path.Combine(TestContext.TestDeploymentDir, $"{TestContext.TestName}.xlsx");
            excelPackage.SaveAs(new FileInfo(savePath));

            var exApp = new Microsoft.Office.Interop.Excel.Application();

            try
            {
                var exWbk = exApp.Workbooks.Open(savePath);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Assert.Fail("It is not possible to open the workbook after EPPlus saved it.");
            }
            finally
            {
                exApp.Workbooks.Close();
            }
#endif
        }

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
        public void Test_rounded_values()
        {
            //Arrange
#if Core
            var dir = AppContext.BaseDirectory;
            dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
            var excelPackage =
                new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "TestDoc_WithRoundedValues_xlsx.xlsx")));

            //Act
            excelPackage.Workbook.Calculate();
            var table = excelPackage.Workbook.Worksheets[0];

            var value1 = table.Cells["A1"].Value.ToString();
            var value2 = table.Cells["A4"].Value.ToString();
            var value3 = table.Cells["B4"].Value.ToString();

            //Asserts
            Assert.IsTrue(value1.Equals("-18"));
            Assert.IsTrue(value2.Equals("-40,5"));
            Assert.IsTrue(value3.Equals("-23,4"));
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
        public void Workbook_Styles()
        {
            // ARRANGE
            var xlsx = GetTestStream("Layout_Format_vorlage.xlsx");
            var package = new ExcelPackage(xlsx);
            
            // ACT
            var styles = package.Workbook.Styles;
            
            // ASSERT
            Assert.AreEqual(0, styles.CellStyleXfs[0].NumberFormatId);
            Assert.AreEqual(0, styles.CellStyleXfs[0].FontId);
            Assert.AreEqual(0, styles.CellStyleXfs[0].FillId);
            Assert.AreEqual(0, styles.CellStyleXfs[0].BorderId);
            Assert.IsNull(styles.CellStyleXfs[0].ApplyNumberFormat);
            Assert.IsNull(styles.CellStyleXfs[0].ApplyFill);
            Assert.IsNull(styles.CellStyleXfs[0].ApplyBorder);
            Assert.IsNull(styles.CellStyleXfs[0].ApplyAlignment);
            Assert.IsNull(styles.CellStyleXfs[0].ApplyProtection);

            Assert.AreEqual(0, styles.CellStyleXfs[1].NumberFormatId);
            Assert.AreEqual(1, styles.CellStyleXfs[1].FontId);
            Assert.AreEqual(0, styles.CellStyleXfs[1].FillId);
            Assert.AreEqual(0, styles.CellStyleXfs[1].BorderId);
            Assert.AreEqual(false, styles.CellStyleXfs[1].ApplyNumberFormat);
            Assert.AreEqual(false, styles.CellStyleXfs[1].ApplyFill);
            Assert.AreEqual(false, styles.CellStyleXfs[1].ApplyBorder);
            Assert.AreEqual(false, styles.CellStyleXfs[1].ApplyAlignment);
            Assert.AreEqual(false, styles.CellStyleXfs[1].ApplyProtection);

            Assert.AreEqual(0, styles.CellXfs[0].NumberFormatId);
            Assert.AreEqual(0, styles.CellXfs[0].FontId);
            Assert.AreEqual(0, styles.CellXfs[0].FillId);
            Assert.AreEqual(0, styles.CellXfs[0].BorderId);
            Assert.IsNull(styles.CellXfs[0].ApplyNumberFormat);
            Assert.IsNull(styles.CellXfs[0].ApplyFill);
            Assert.IsNull(styles.CellXfs[0].ApplyBorder);
            Assert.IsNull(styles.CellXfs[0].ApplyAlignment);
            Assert.IsNull(styles.CellXfs[0].ApplyProtection);

            Assert.AreEqual(0, styles.CellXfs[1].NumberFormatId);
            Assert.AreEqual(1, styles.CellXfs[1].FontId);
            Assert.AreEqual(0, styles.CellXfs[1].FillId);
            Assert.AreEqual(0, styles.CellXfs[1].BorderId);
            Assert.IsNull(styles.CellXfs[1].ApplyNumberFormat);
            Assert.IsNull(styles.CellXfs[1].ApplyFill);
            Assert.IsNull(styles.CellXfs[1].ApplyBorder);
            Assert.IsNull(styles.CellXfs[1].ApplyAlignment);
            Assert.IsNull(styles.CellXfs[1].ApplyProtection);
        }

        [TestMethod]
        public void ExcelPackage_Calculate()
        {
            // Arrange
            var input = GetTestStream("Trim.xlsx");
            var package = new ExcelPackage(input);
            var sheet = package.Workbook.Worksheets[1];

            // Act
            sheet.Calculate();

            // Arrange
            Assert.AreEqual("Anlagevermögen", sheet.Cells["B8"].Value);
            Assert.AreEqual("123 456 ABC", sheet.Cells["B9"].Value);
        }
        [TestMethod]
        public void Test_FormulaSumPrecision()
        {

            // Arrange
            var input = GetTestStream("TestDoc_SumPrecision.xlsx");
            var package = new ExcelPackage(input);
            var ws = package.Workbook.Worksheets["Tabelle1"];

            // Act
            ws.Calculate();

            //Arrange
#if Core
            var dir = AppContext.BaseDirectory;
            dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif

            var result = ws.Cells["L14"].Value.ToString();
            Assert.AreEqual("-3,552713678800501E-15", result);


           
        }
    }
}