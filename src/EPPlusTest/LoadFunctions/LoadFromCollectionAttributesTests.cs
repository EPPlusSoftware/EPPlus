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
    [EpplusTable(TableStyle = TableStyles.Dark1, PrintHeaders = true, AutofitColumns = true, AutoCalculate = true, ShowTotal = true, ShowFirstColumn = true)]
    [
        EpplusFormulaTableColumn(Order = 6, NumberFormat = "€#,##0.00", Header = "Tax amount", FormulaR1C1 = "RC[-2] * RC[-1]", TotalsRowFunction = RowFunctions.Sum, TotalsRowNumberFormat = "€#,##0.00"),
        EpplusFormulaTableColumn(Order = 7, NumberFormat = "€#,##0.00", Header = "Net salary", Formula = "E2-G2", TotalsRowFunction = RowFunctions.Sum, TotalsRowNumberFormat = "€#,##0.00")
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

        [EpplusTableColumn(Order = 0, NumberFormat = "yyyy-MM-dd", TotalsRowLabel = "Total")]
        public DateTime Birthdate { get; set; }

        [EpplusTableColumn(Order = 4, NumberFormat = "€#,##0.00", TotalsRowFunction = RowFunctions.Sum, TotalsRowNumberFormat = "€#,##0.00")]
        public double Salary { get; set; }

        [EpplusTableColumn(Order = 5, NumberFormat = "0%", TotalsRowFormula = "Table1[[#Totals],[Tax amount]]/Table1[[#Totals],[Salary]]", TotalsRowNumberFormat ="0 %")]
        public double Tax { get; set; }
    }

    [EpplusTable(TableStyle = TableStyles.Medium1, PrintHeaders = true, AutofitColumns = true, AutoCalculate = true, ShowLastColumn = true)]
    internal class Actor2 : Actor
    {

    }

    [EpplusTable(TableStyle = TableStyles.None, PrintHeaders = true, AutofitColumns = true, AutoCalculate = true, ShowLastColumn = true)]
    internal class ActorTablestyleNone : Actor
    {

    }

    [TestClass]
    public class LoadFromCollectionAttributesTests
    {
        private readonly List<Actor> _actors = new List<Actor>
        {
            new Actor{ Salary = 256.24, Tax = 0.21, FirstName = "John", MiddleName = "Bernhard", LastName = "Doe", Birthdate = new DateTime(1950, 3, 15) },
            new Actor{ Salary = 278.55, Tax = 0.23, FirstName = "Sven", MiddleName = "Bertil", LastName = "Svensson", Birthdate = new DateTime(1962, 6, 10)},
            new Actor{ Salary = 315.34, Tax = 0.28, FirstName = "Lisa", MiddleName = "Maria", LastName = "Gonzales", Birthdate = new DateTime(1971, 10, 2)}
        };

        [TestMethod]
        public void ShouldUseAttributeForSorting()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromCollection(_actors);

                Assert.AreEqual("Birthdate", sheet.Cells["A1"].Value);
                Assert.AreEqual("First name", sheet.Cells["B1"].Value);
                Assert.AreEqual("Tax", sheet.Cells["F1"].Value);
                Assert.AreEqual("John", sheet.Cells["B2"].Value);
                Assert.AreEqual("Svensson", sheet.Cells["D3"].Value);
                Assert.AreEqual(0.28, sheet.Cells["F4"].Value);

                //package.SaveAs(new FileInfo(@"c:\temp\coll.xlsx"));
            }
        }

        [TestMethod]
        public void ShouldUseAttributeForTableStyle()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromCollection(_actors);
                var table = sheet.Tables[0];
                Assert.AreEqual(TableStyles.Dark1, table.TableStyle);
            }
        }

        public void ShouldNotAutoCalc()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromCollection(_actors);
                Assert.IsNull(sheet.Cells["H3"].Value);
            }
        }

        [TestMethod]
        public void InheritedShouldAutoCalc()
        {
            var actors = new List<Actor2>
            {
                new Actor2{ Salary = 256.24, Tax = 0.21, FirstName = "John", MiddleName = "Bernhard", LastName = "Doe", Birthdate = new DateTime(1950, 3, 15) },
                new Actor2{ Salary = 278.55, Tax = 0.23, FirstName = "Sven", MiddleName = "Bertil", LastName = "Svensson", Birthdate = new DateTime(1962, 6, 10)},
                new Actor2{ Salary = 315.34, Tax = 0.28, FirstName = "Lisa", MiddleName = "Maria", LastName = "Gonzales", Birthdate = new DateTime(1971, 10, 2)}
            };
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromCollection(actors);
                var table = sheet.Tables[0];
                Assert.AreEqual(TableStyles.Medium1, table.TableStyle);
                Assert.IsNotNull(sheet.Cells["H3"].Value);
            }
        }

        [TestMethod]
        public void TableStyleNoneShouldAutoCalc()
        {
            var actors = new List<ActorTablestyleNone>
            {
                new ActorTablestyleNone{ Salary = 256.24, Tax = 0.21, FirstName = "John", MiddleName = "Bernhard", LastName = "Doe", Birthdate = new DateTime(1950, 3, 15) },
                new ActorTablestyleNone{ Salary = 278.55, Tax = 0.23, FirstName = "Sven", MiddleName = "Bertil", LastName = "Svensson", Birthdate = new DateTime(1962, 6, 10)},
                new ActorTablestyleNone{ Salary = 315.34, Tax = 0.28, FirstName = "Lisa", MiddleName = "Maria", LastName = "Gonzales", Birthdate = new DateTime(1971, 10, 2)}
            };
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromCollection(actors);
                var table = sheet.Tables[0];
                Assert.AreEqual(TableStyles.None, table.TableStyle);
                Assert.IsNotNull(sheet.Cells["H3"].Value);
            }
        }

        [TestMethod]
        public void ShouldUseFuncArgOverAttributesForHeaders()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromCollection(_actors, false);

                Assert.AreEqual("John", sheet.Cells["B1"].Value);
                Assert.AreEqual("Svensson", sheet.Cells["D2"].Value);
                Assert.AreEqual(0.28, sheet.Cells["F3"].Value);
            }
        }

        [TestMethod]
        public void ShouldUseFuncArgOverAttributeForTableStyle()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromCollection(_actors, true, TableStyles.Dark4);
                var table = sheet.Tables[0];
                Assert.AreEqual(TableStyles.Dark4, table.TableStyle);
            }
        }

        [TestMethod]
        public void ShouldUseSortOrderAttributeOnClassLevel()
        {
            var objects = new OuterWithSortOrderOnClassLevelV1[]
            {
                new OuterWithSortOrderOnClassLevelV1{ ApprovedUtc = new DateTime(2021, 12, 14), Acknowledged = true, Organization = new Organization()},
                new OuterWithSortOrderOnClassLevelV1{ ApprovedUtc = new DateTime(2021, 12, 15), Acknowledged = false, Organization = new Organization()}
            };
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var r = sheet.Cells["A1"].LoadFromCollection(objects, true, TableStyles.Dark4);
                var table = sheet.Tables[0];
                Assert.AreEqual("Acknowledged...", sheet.Cells["A1"].Value);
                Assert.AreEqual("Org Level 4", sheet.Cells["B1"].Value);
                Assert.AreEqual("ApprovedUtc", sheet.Cells["C1"].Value);

            }
        }
    }
}
