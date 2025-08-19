using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace EPPlusTest.LoadFunctions
{
    [TestClass]
    public class LoadFromCollectionAttributesMemberFilterTests
    {
        [TestMethod]
        public void ShouldFilterNestedPropertiesByMemberList()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var items = new List<LfcaTestClass1>
                {
                    new LfcaTestClass1{ Id = 1, Item = new LfcaTestClass2{ Id = 2, Name = "Test 1"}},
                    new LfcaTestClass1{ Id = 3, Item = new LfcaTestClass2{ Id = 4, Name = "Test 1"}}
                };
                var t = typeof(LfcaTestClass1);
                var t2 = typeof(LfcaTestClass2);
                sheet.Cells["A1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                    c.Members = new MemberInfo[]
                    {
                        t.GetProperty("Id"),
                        t2.GetProperty("Id"),
                        t2.GetProperty("Name")
                    };
                });

                Assert.AreEqual("Id", sheet.Cells["A1"].Value);
                Assert.AreEqual("Class 2 Id", sheet.Cells["B1"].Value);
                Assert.AreEqual("Class 2 Name", sheet.Cells["C1"].Value);
                Assert.AreEqual(1, sheet.Cells["A2"].Value);
                Assert.AreEqual(2, sheet.Cells["B2"].Value);
                Assert.AreEqual("Test 1", sheet.Cells["C2"].Value);
                Assert.IsNull(sheet.Cells["D1"].Value);
            }
        }

        [TestMethod]
        public void ShouldFilterNestedPropertiesByMemberListNested()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var items = new List<LfcaTestClass1>
                {
                    new LfcaTestClass1{ Id = 1, Item = new LfcaTestClass2{ Id = 2, Name = "Test 1"}},
                    new LfcaTestClass1{ Id = 3, Item = new LfcaTestClass2{ Id = 4, Name = "Test 1"}}
                };
                var t = typeof(LfcaTestClass1);
                var t2 = typeof(LfcaTestClass2);
                sheet.Cells["A1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                    c.Members = new MemberInfo[]
                    {
                        t.GetProperty("Id"),
                        t.GetProperty("Item"),
                        t2.GetProperty("Name"),
                        t2.GetProperty("Id")
                    };
                });

                Assert.AreEqual("Id", sheet.Cells["A1"].Value);
                Assert.AreEqual("Class 2 Name", sheet.Cells["B1"].Value);
                Assert.AreEqual("Class 2 Id", sheet.Cells["C1"].Value);
            }
        }

        [TestMethod]
        public void ShouldFilterNestedPropertiesByMemberList2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                var items = new List<LfcaTestClass1>
                {
                    new LfcaTestClass1{ Id = 1, Item2 = new LfcaTestClass3{ Id = 2, Name = "Test 1"}},
                    new LfcaTestClass1{ Id = 3, Item2 = new LfcaTestClass3{ Id = 4, Name = "Test 1"}}
                };
                var t = typeof(LfcaTestClass1);
                var t2 = typeof(LfcaTestClass2);
                sheet.Cells["A1"].LoadFromCollection(items, c =>
                {
                    c.PrintHeaders = true;
                    c.Members = new MemberInfo[]
                    {
                        t.GetProperty("Id"),
                        t.GetProperty("Item2")
                    };
                });

                Assert.AreEqual("Id", sheet.Cells["A1"].Value);
                Assert.AreEqual("Class 3 Name", sheet.Cells["B1"].Value);
                Assert.AreEqual(1, sheet.Cells["A2"].Value);
                Assert.AreEqual("Test 1", sheet.Cells["B2"].Value);
                Assert.IsNull(sheet.Cells["C1"].Value);
            }
        }
    }

    #region Test classes
    [EpplusTable(TableStyle = OfficeOpenXml.Table.TableStyles.Light1)]
    internal class LfcaTestClass1
    {
        [EpplusTableColumn(Order = 1)]
        public int Id { get; set; }

        [EpplusNestedTableColumn(HeaderPrefix = "Class 2", Order = 2)]
        public LfcaTestClass2 Item { get; set; }

        [EpplusNestedTableColumn(HeaderPrefix = "Class 3", Order = 3)]
        public LfcaTestClass3 Item2 { get; set; }
    }

    internal class LfcaTestClass2
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }

    internal class LfcaTestClass3
    {
        [EpplusIgnore]
        public int Id { get; set; }
        public string Name { get; set; }
    }
    #endregion
}
