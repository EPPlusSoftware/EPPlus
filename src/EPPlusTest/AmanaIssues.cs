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
    }
}