using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.LoadFunctions
{
    [TestClass]
    public class LoadFromCollectionSortOrderAttributeTests
    {
        [EPPlusTableColumnSortOrder(Properties = new string[] { nameof(Name), nameof(Obj), nameof(Id) })]
        [EpplusTable]
        public class Outer
        {
            [EpplusTableColumn]
            public int Id { get; set; }

            [EpplusNestedTableColumn]
            public Inner Obj { get; set; }

            [EpplusTableColumn]
            public string Name { get; set; }
        }

        [EPPlusTableColumnSortOrder(Properties = new string[] { nameof(Name), nameof(Id) })]
        public class Inner
        {
            [EpplusTableColumn]
            public int Id { get; set; }

            [EpplusTableColumn]
            public string Name { get; set; }
        }

        [EPPlusTableColumnSortOrder(Properties = new string[] { nameof(Name), nameof(Obj), nameof(Id) })]
        [EpplusTable]
        public class Outer2
        {
            [EpplusTableColumn]
            public int Id { get; set; }

            [EpplusNestedTableColumn]
            public Inner2 Obj { get; set; }

            [EpplusTableColumn]
            public string Name { get; set; }
        }

        [EPPlusTableColumnSortOrder(Properties = new string[] { nameof(Name), nameof(Id), nameof(Email) })]
        public class Inner2
        {
            [EpplusTableColumn]
            public int Id { get; set; }

            [EpplusTableColumn]
            public string Name { get; set; }
            
            [EpplusNestedTableColumn]
            public InnerEmail Email { get; set; }
        }

        [EPPlusTableColumnSortOrder(Properties = new string[] { nameof(Email), nameof(Id) })]
        public class InnerEmail
        {
            [EpplusTableColumn]
            public int Id { get; set; }

            [EpplusTableColumn]
            public string Email { get; set; }
        }

        [TestMethod]
        public void SortBySortorderAttribute1()
        {
            var outer = new Outer
            {
                Id = 1,
                Obj = new Inner
                {
                    Id = 2,
                    Name = "Inner"
                },
                Name = "Outer"
            };
            var items = new List<Outer> { outer };
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet 1");
                sheet.Cells["A1"].LoadFromCollection(items);

                Assert.AreEqual("Outer", sheet.Cells["A1"].Value);
                Assert.AreEqual("Inner", sheet.Cells["B1"].Value);
                Assert.AreEqual(2, sheet.Cells["C1"].Value);
                Assert.AreEqual(1, sheet.Cells["D1"].Value);
            }
        }

        [TestMethod]
        public void SortBySortorderAttribute2()
        {
            var outer = new Outer2
            {
                Id = 1,
                Obj = new Inner2
                {
                    Id = 2,
                    Name = "Inner",
                    Email = new InnerEmail
                    {
                        Id = 3,
                        Email = "inner@outer.com"
                    }
                },
                Name = "Outer"
            };
            var items = new List<Outer2> { outer };
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet 1");
                sheet.Cells["A1"].LoadFromCollection(items);

                Assert.AreEqual("Outer", sheet.Cells["A1"].Value);
                Assert.AreEqual("Inner", sheet.Cells["B1"].Value);
                Assert.AreEqual(2, sheet.Cells["C1"].Value);
                Assert.AreEqual("inner@outer.com", sheet.Cells["D1"].Value);
                Assert.AreEqual(3, sheet.Cells["E1"].Value);
                Assert.AreEqual(1, sheet.Cells["F1"].Value);
            }
        }
    }
}
