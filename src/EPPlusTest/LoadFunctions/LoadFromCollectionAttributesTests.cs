using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.LoadFunctions
{
    [EpplusTable(TableStyle = TableStyles.Dark1, PrintHeaders = true, AutofitColumns = true, AutoCalculate = false)]
    [
        EpplusFormulaTableColumn(ColumnOrder = 6, ColumnNumberFormat = "€#,##0.00", ColumnHeader = "Tax amount", Formula = "E{row} * F{row}"),
        EpplusFormulaTableColumn(ColumnOrder = 7, ColumnNumberFormat = "€#,##0.00", ColumnHeader = "Net salary", Formula = "E{row} - G{row}")
    ]
    internal class Composer
    {
        [EpplusIgnore]
        public int Id { get; set; }

        [EpplusTableColumn(ColumnOrder = 3)]
        public string LastName { get; set; }
        [EpplusTableColumn(ColumnOrder = 1, ColumnHeader = "First name")]
        public string FirstName { get; set; }
        [EpplusTableColumn(ColumnOrder = 2)]
        public string MiddleName { get; set; }

        [EpplusTableColumn(ColumnOrder = 0, ColumnNumberFormat = "yyyy-MM-dd")]
        public DateTime Timestamp { get; set; }

        [EpplusTableColumn(ColumnOrder = 4, ColumnNumberFormat = "€#,##0.00")]
        public double Salary { get; set; }

        [EpplusTableColumn(ColumnOrder = 5, ColumnNumberFormat = "0%")]
        public double Tax { get; set; }
    }

    [TestClass]
    public class LoadFromCollectionAttributesTests
    {

        [TestMethod]
        public void ShouldUseAttributeForSorting()
        {
            var items = new List<Composer>
            {
                new Composer{ Salary = 256.24, Tax = 0.21, FirstName = "Johann", MiddleName = "Sebastian", LastName = "Bach", Timestamp = DateTime.Now },
                new Composer{ Salary = 278.55, Tax = 0.23, FirstName = "Wolfgang", MiddleName = "Amadeus", LastName = "Mozart", Timestamp = DateTime.Now.AddDays(1)}
            };
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromCollection(items);

                Assert.AreEqual("First name", sheet.Cells["B1"].Value);
                Assert.AreEqual("Johann", sheet.Cells["B2"].Value);
                Assert.AreEqual("Mozart", sheet.Cells["D3"].Value);

                //package.SaveAs(new FileInfo(@"c:\temp\coll.xlsx"));
            }
        }
    }
}
