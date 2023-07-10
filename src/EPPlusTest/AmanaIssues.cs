namespace EPPlusTest
{
    using EPPlusTest.Properties;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using OfficeOpenXml;
    using System;
    using System.IO;


    [TestClass]
    public class AmanaIssues : TestBase
    {

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
    }
}