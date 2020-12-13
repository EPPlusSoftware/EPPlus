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
        EpplusFormulaTableColumn(Order = 6, NumberFormat = "€#,##0.00", Header = "Tax amount", FormulaR1C1 = "RC[-2] * RC[-1]"),
        EpplusFormulaTableColumn(Order = 7, NumberFormat = "€#,##0.00", Header = "Net salary", Formula = "E2-G2")
    ]
    internal class Actor
    {
        [EpplusIgnore]
        public int Id { get; set; }

        [EpplusTableColumn(Order = 3)]
        public string LastName { get; set; }
        [EpplusTableColumn(Order = 1, Header = "First name")]
        public string FirstName { get; set; }
        [EpplusTableColumn(Order = 2)]
        public string MiddleName { get; set; }

        [EpplusTableColumn(Order = 0, NumberFormat = "yyyy-MM-dd")]
        public DateTime Birthdate { get; set; }

        [EpplusTableColumn(Order = 4, NumberFormat = "€#,##0.00")]
        public double Salary { get; set; }

        [EpplusTableColumn(Order = 5, NumberFormat = "0%")]
        public double Tax { get; set; }
    }

    [TestClass]
    public class LoadFromCollectionAttributesTests
    {

        [TestMethod]
        public void ShouldUseAttributeForSorting()
        {
            var items = new List<Actor>
            {
                new Actor{ Salary = 256.24, Tax = 0.21, FirstName = "John", MiddleName = "Bernhard", LastName = "Doe", Birthdate = new DateTime(1950, 3, 15) },
                new Actor{ Salary = 278.55, Tax = 0.23, FirstName = "Sven", MiddleName = "Bertil", LastName = "Svensson", Birthdate = new DateTime(1962, 6, 10)},
                new Actor{ Salary = 315.34, Tax = 0.28, FirstName = "Lisa", MiddleName = "Maria", LastName = "Gonzales", Birthdate = new DateTime(1971, 10, 2)}
            };
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromCollection(items);

                Assert.AreEqual("First name", sheet.Cells["B1"].Value);
                Assert.AreEqual("John", sheet.Cells["B2"].Value);
                Assert.AreEqual("Svensson", sheet.Cells["D3"].Value);

                //package.SaveAs(new FileInfo(@"c:\temp\coll.xlsx"));
            }
        }
    }
}
