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
            var excelPackage =
                new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "TestDoc_XMLTagsTable_xlsx.xlsx")));

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
    }
}