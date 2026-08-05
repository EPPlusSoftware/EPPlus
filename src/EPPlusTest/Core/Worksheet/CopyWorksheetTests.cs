using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;

namespace EPPlusTest.Core.Worksheet
{
    [TestClass]
    public class CopyWorksheetTests : TestBase
    {
        private static ExcelPackage CreatePackageWithTable(out ExcelWorksheet source)
        {
            var package = new ExcelPackage();
            source = package.Workbook.Worksheets.Add("Template");
            source.Cells["A1"].Value = "Header";
            source.Cells["A2"].Value = 1;
            source.Cells["A3"].Value = 2;
            source.Tables.Add(source.Cells["A1:A3"], "BoxOffice");
            return package;
        }

        [TestMethod]
        public void Copy_WithTableCopyHandler_RenamesCopiedTable()
        {
            using (var package = CreatePackageWithTable(out var source))
            {
                var copy = package.Workbook.Worksheets.Copy(source.Name, "BaltimoreMD", options =>
                {
                    options.TableCopyHandler = args =>
                    {
                        args.NewName = "BaltimoreMD_" + args.SourceTableName;
                    };
                });

                Assert.AreEqual(1, copy.Tables.Count);
                Assert.IsNotNull(copy.Tables["BaltimoreMD_BoxOffice"]);
            }
        }

        [TestMethod]
        public void Copy_WithTableCopyHandler_ProvidesSourceAndDefaultName()
        {
            using (var package = CreatePackageWithTable(out var source))
            {
                string capturedSourceName = null;
                string capturedDefaultName = null;

                package.Workbook.Worksheets.Copy(source.Name, "Copy", options =>
                {
                    options.TableCopyHandler = args =>
                    {
                        capturedSourceName = args.SourceTableName;
                        capturedDefaultName = args.DefaultName;
                    };
                });

                Assert.AreEqual("BoxOffice", capturedSourceName);
                // Same workbook, so the copy is given a generated name.
                Assert.IsTrue(capturedDefaultName.StartsWith("Table"));
            }
        }

        [TestMethod]
        public void Copy_TableCopyHandler_UpdatesFormulaReferencesToRenamedTable()
        {
            using (var package = CreatePackageWithTable(out var source))
            {
                // A cell on the source that references the table by name.
                source.Cells["C1"].Formula = "SUM(BoxOffice[Header])";

                var copy = package.Workbook.Worksheets.Copy(source.Name, "Renamed", options =>
                {
                    options.TableCopyHandler = args =>
                    {
                        args.NewName = "NewBoxOffice";
                    };
                });

                // The copied formula should now reference the renamed table,
                // updated token based through the ExcelTable.Name setter.
                Assert.AreEqual("SUM(NewBoxOffice[Header])", copy.Cells["C1"].Formula);
            }
        }

        [TestMethod]
        public void Copy_TableCopyHandler_NullNewName_KeepsDefaultName()
        {
            using (var package = CreatePackageWithTable(out var source))
            {
                string defaultName = null;

                var copy = package.Workbook.Worksheets.Copy(source.Name, "Copy", options =>
                {
                    options.TableCopyHandler = args =>
                    {
                        defaultName = args.DefaultName;
                        // NewName left null.
                    };
                });

                Assert.IsNotNull(copy.Tables[defaultName]);
            }
        }

        [TestMethod]
        public void Copy_WithoutOptions_KeepsExistingBehavior()
        {
            using (var package = CreatePackageWithTable(out var source))
            {
                var copy = package.Workbook.Worksheets.Copy(source.Name, "Copy");

                Assert.AreEqual(1, copy.Tables.Count);
                // Same workbook copy generates a Table{n} name.
                Assert.IsTrue(copy.Tables[0].Name.StartsWith("Table"));
            }
        }

        [TestMethod]
        public void Copy_TableCopyHandler_RenameToExistingName_Throws()
        {
            using (var package = CreatePackageWithTable(out var source))
            {
                // A second table whose name we will collide with.
                source.Cells["E1"].Value = "H";
                source.Cells["E2"].Value = 1;
                source.Tables.Add(source.Cells["E1:E2"], "Ancillary");

                Assert.ThrowsExactly<ArgumentException>(() =>
                {
                    package.Workbook.Worksheets.Copy(source.Name, "Copy", options =>
                    {
                        options.TableCopyHandler = args =>
                        {
                            // Force both copied tables to the same name.
                            args.NewName = "Duplicate";
                        };
                    });
                });
            }
        }

        [TestMethod]
        public void Copy_MultipleTables_RenamesOnlySelectedSubset()
        {
            using (var package = CreatePackageWithTable(out var source))
            {
                source.Cells["E1"].Value = "H";
                source.Cells["E2"].Value = 1;
                source.Cells["E3"].Value = 2;
                source.Tables.Add(source.Cells["E1:E3"], "Ancillary");

                var copy = package.Workbook.Worksheets.Copy(source.Name, "Copy", options =>
                {
                    options.TableCopyHandler = args =>
                    {
                        if (args.SourceTableName == "BoxOffice")
                        {
                            args.NewName = "Copied_BoxOffice";
                        }
                        // Ancillary left with its default name.
                    };
                });

                Assert.IsNotNull(copy.Tables["Copied_BoxOffice"]);
                Assert.AreEqual(2, copy.Tables.Count);
            }
        }

        [TestMethod]
        public void Copy_TableCopyHandler_DoesNotAffectSourceWorksheetFormulas()
        {
            using (var package = CreatePackageWithTable(out var source))
            {
                source.Cells["C1"].Formula = "SUM(BoxOffice[Header])";

                package.Workbook.Worksheets.Copy(source.Name, "Renamed", options =>
                {
                    options.TableCopyHandler = args =>
                    {
                        args.NewName = "NewBoxOffice";
                    };
                });

                //The source worksheet still has its own BoxOffice table; its formula must be untouched.
                Assert.AreEqual("SUM(BoxOffice[Header])", source.Cells["C1"].Formula);
            }
        }

        [TestMethod]
        public void Copy_WithoutHandler_AssignsGeneratedTableName()
        {
            using (var package = CreatePackageWithTable(out var source))
            {
                var copy = package.Workbook.Worksheets.Copy(source.Name, "Copy");

                Assert.AreEqual(1, copy.Tables.Count);
                //Same workbook copy renames the copied table to a generated Table{n} name.
                Assert.AreNotEqual("BoxOffice", copy.Tables[0].Name);
                Assert.IsTrue(copy.Tables[0].Name.StartsWith("Table"));
            }
        }

        [TestMethod]
        public void Copy_WithoutHandler_GeneratedNameIsUniqueInWorkbook()
        {
            using (var package = CreatePackageWithTable(out var source))
            {
                var copy = package.Workbook.Worksheets.Copy(source.Name, "Copy");

                //Source keeps its original name, the copy gets a distinct generated name.
                Assert.AreEqual("BoxOffice", source.Tables[0].Name);
                Assert.AreNotEqual(source.Tables[0].Name, copy.Tables[0].Name);
                Assert.IsFalse(source.Tables[0].Name.Equals(copy.Tables[0].Name, StringComparison.OrdinalIgnoreCase));
            }
        }

        [TestMethod]
        public void Copy_WithoutHandler_AdjustsCopiedFormulaToGeneratedName()
        {
            using (var package = CreatePackageWithTable(out var source))
            {
                source.Cells["C1"].Formula = "SUM(BoxOffice[Header])";

                var copy = package.Workbook.Worksheets.Copy(source.Name, "Copy");

                //The copied formula must reference the copied table's generated name,
                //not the source table name.
                var expected = "SUM(" + copy.Tables[0].Name + "[Header])";
                Assert.AreEqual(expected, copy.Cells["C1"].Formula);
            }
        }

        [TestMethod]
        public void Copy_WithoutHandler_DoesNotAffectSourceFormula()
        {
            using (var package = CreatePackageWithTable(out var source))
            {
                source.Cells["C1"].Formula = "SUM(BoxOffice[Header])";

                package.Workbook.Worksheets.Copy(source.Name, "Copy");

                //Source formula still references the source table, untouched by the copy.
                Assert.AreEqual("SUM(BoxOffice[Header])", source.Cells["C1"].Formula);
            }
        }

        [TestMethod]
        public void Copy_WithoutHandler_MultipleTables_AllGetGeneratedNames()
        {
            using (var package = CreatePackageWithTable(out var source))
            {
                source.Cells["E1"].Value = "H";
                source.Cells["E2"].Value = 1;
                source.Cells["E3"].Value = 2;
                source.Tables.Add(source.Cells["E1:E3"], "Ancillary");

                var copy = package.Workbook.Worksheets.Copy(source.Name, "Copy");

                Assert.AreEqual(2, copy.Tables.Count);
                foreach (var t in copy.Tables)
                {
                    Assert.IsTrue(t.Name.StartsWith("Table"));
                }
                //The two copied tables get distinct names.
                Assert.AreNotEqual(copy.Tables[0].Name, copy.Tables[1].Name);
            }
        }

        [TestMethod]
        public void Copy_WithoutHandler_MultipleTables_AdjustsEachFormulaToItsOwnCopy()
        {
            using (var package = CreatePackageWithTable(out var source))
            {
                source.Cells["E1"].Value = "H";
                source.Cells["E2"].Value = 1;
                source.Cells["E3"].Value = 2;
                source.Tables.Add(source.Cells["E1:E3"], "Ancillary");

                source.Cells["G1"].Formula = "SUM(BoxOffice[Header])";
                source.Cells["G2"].Formula = "SUM(Ancillary[H])";

                var copy = package.Workbook.Worksheets.Copy(source.Name, "Copy");

                //Resolve which copied table owns which column by header, since names are generated.
                string boxOfficeCopyName = null;
                string ancillaryCopyName = null;
                foreach (var t in copy.Tables)
                {
                    if (t.Columns[0].Name == "Header") boxOfficeCopyName = t.Name;
                    if (t.Columns[0].Name == "H") ancillaryCopyName = t.Name;
                }

                Assert.AreEqual("SUM(" + boxOfficeCopyName + "[Header])", copy.Cells["G1"].Formula);
                Assert.AreEqual("SUM(" + ancillaryCopyName + "[H])", copy.Cells["G2"].Formula);
            }
        }
    }
}