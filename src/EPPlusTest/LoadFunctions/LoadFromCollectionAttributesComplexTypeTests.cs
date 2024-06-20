using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.LoadFunctions;
using OfficeOpenXml.LoadFunctions.Params;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Table;

namespace EPPlusTest.LoadFunctions
{
    [TestClass]
    public class LoadFromCollectionAttributesComplexTypeTests : TestBase
    {
        private List<Outer> _collection = new List<Outer>();
        private List<OuterWithHeaders> _collectionHeaders = new List<OuterWithHeaders>();
        private List<OuterReversedSortOrder> _collectionReversed = new List<OuterReversedSortOrder>();
        private List<OuterSubclass> _collectionInheritence = new List<OuterSubclass>();
        private List<ColumnsWithoutAttributes> _collectionNoAttributes = new List<ColumnsWithoutAttributes>();

        [TestInitialize]
        public void Initialize()
        {
            _collection.Add(new Outer
            {
                ApprovedUtc = new DateTime(2021, 7, 1),
                Organization = new Organization 
                { 
                    OrgLevel3 = "ABC", 
                    OrgLevel4 = "DEF", 
                    OrgLevel5 = "GHI"
                },
                Acknowledged = true
            });
            _collectionHeaders.Add(new OuterWithHeaders
            {
                ApprovedUtc = new DateTime(2021, 7, 1),
                Organization = new Organization
                {
                    OrgLevel3 = "ABC",
                    OrgLevel4 = "DEF",
                    OrgLevel5 = "GHI"
                },
                Acknowledged = true
            });
            _collectionReversed.Add(new OuterReversedSortOrder
            {
                ApprovedUtc = new DateTime(2021, 7, 1),
                Organization = new OrganizationReversedSortOrder
                {
                    OrgLevel3 = "ABC",
                    OrgLevel4 = "DEF",
                    OrgLevel5 = "GHI"
                },
                Acknowledged = true
            });
            _collectionInheritence.Add(new OuterSubclass
            {
                ApprovedUtc = new DateTime(2021, 7, 1),
                Organization = new OrganizationSubclass
                {
                    OrgLevel3 = "ABC",
                    OrgLevel4 = "DEF",
                    OrgLevel5 = "GHI"
                },
                Acknowledged = true
            });
            _collectionNoAttributes.Add(new ColumnsWithoutAttributes
            {
                NullableInt = 5,
                NonNull = 15,
                NullableDateTime = new DateTime(2021, 7, 1),
                NestedNullableNullable = new NestedNullable { NullableValue = -2 },
                ExplicitlyNullableString = "I'm nullable"
            }) ;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _collection.Clear();
        }

        [TestMethod]
        public void ShouldSetupColumnsWithPath()
        {
            var parameters = new LoadFromCollectionParams
            {
                BindingFlags = LoadFromCollectionParams.DefaultBindingFlags
            };
            var cols = new LoadFromCollectionColumns<Outer>(parameters);
            var result = cols.Setup();
            Assert.AreEqual(5, result.Count, "List did not contain 5 elements as expected");
            Assert.AreEqual("ApprovedUtc", result[0].Path.GetPath());
            Assert.AreEqual("Organization.OrgLevel3", result[1].Path.GetPath());
        }

        [TestMethod]
        public void ShouldSetupColumnsWithPathSorted()
        {
            var parameters = new LoadFromCollectionParams
            {
                BindingFlags = LoadFromCollectionParams.DefaultBindingFlags
            };
            var cols = new LoadFromCollectionColumns<OuterReversedSortOrder>(parameters);
            var result = cols.Setup();
            Assert.AreEqual(5, result.Count, "List did not contain 5 elements as expected");
            Assert.AreEqual("Acknowledged", result[0].Path.GetPath());
            Assert.AreEqual("Organization.OrgLevel5", result[1].Path.GetPath());
            Assert.AreEqual("ApprovedUtc", result.Last().Path.GetPath());
        }

        [TestMethod]
        public void ShouldLoadFromComplexTypeMember()
        {
            using(var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].LoadFromCollection(_collection);
                Assert.AreEqual("ABC", ws.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void ShouldLoadFromComplexTypeMemberSorted()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].LoadFromCollection(_collectionReversed);
                Assert.IsTrue((bool)ws.Cells["A1"].Value);
                Assert.AreEqual("GHI", ws.Cells["B1"].Value);
                Assert.AreEqual(new DateTime(2021, 7, 1), ws.Cells["E1"].Value);
            }
        }

        [TestMethod]
        public void ShouldLoadFromComplexTypeMemberWhenComplexMemberIsNull()
        {
            var obj = _collection.First();
            obj.Organization = null;
            _collection[0] = obj;
            using (var package = OpenPackage("testtablcollection.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].LoadFromCollection(_collection);
                Assert.IsNull(ws.Cells["B1"].Value);
                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void ShouldLoadFromComplexTypeMemberWhenComplexMemberIsNull_WithHeaders()
        {
            var obj = _collectionHeaders.First();
            obj.Organization = null;
            _collectionHeaders[0] = obj;
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].LoadFromCollection(_collectionHeaders);
                Assert.AreEqual("Org Level 3", ws.Cells["B1"].Value);
                Assert.IsNull(ws.Cells["B2"].Value);
            }
        }

        [TestMethod]
        public void ShouldSetHeaderPrefixOnComplexClassProperty_WithTableColumnAttributeOnChildProperty()
        {
            var items = ExcelItems.GetItems1();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].LoadFromCollection(items);
                var cv = ws.Cells["G1"].Value;
                Assert.AreEqual("Collateral Owner Email", cv);
            }
        }

        [TestMethod]
        public void ShouldSetHeaderPrefixOnComplexClassProperty_WithoutTableColumnAttributeOnChildProperty()
        {
            var items = ExcelItems.GetItems1();
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].LoadFromCollection(items);
                var cv = ws.Cells["F1"].Value;
                Assert.AreEqual("Collateral Owner Name", cv);
            }
        }

        [TestMethod]
        public void ShouldLoadFromComplexInheritence()
        {
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].LoadFromCollection(_collectionInheritence);
                Assert.AreEqual("ABC", ws.Cells["B1"].Value);
            }
        }

        [TestMethod]
        public void LoadComplexTest2()
        {
            using (var package = new ExcelPackage())
            {
                var items = ExcelItems.GetItems1();
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].LoadFromCollection(items);
                Assert.AreEqual("Product Family", sheet.Cells["A1"].Value);
                Assert.AreEqual("PCH Die Name", sheet.Cells["B1"].Value);
                Assert.AreEqual("Collateral Owner Email", sheet.Cells["G1"].Value);
                Assert.AreEqual("Mission Control Lead Email", sheet.Cells["J1"].Value);
                Assert.AreEqual("Created (GMT)", sheet.Cells["L1"].Value);
            }
        }

        [TestMethod]
        public void HiddenTest1()
        {
            using (var package = new ExcelPackage())
            {
                var items = new List<OuterWithHiddenColumn>
                {
                    new OuterWithHiddenColumn
                    {
                        Active = false,
                        Number = 1,
                        HiddenName = "Hidden 1",
                        Name = "Name 1"
                    },
                    new OuterWithHiddenColumn
                    {
                        Active = true,
                        Number = 2,
                        HiddenName = "Hidden 2",
                        Name = "Name 2"
                    }
                };
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].LoadFromCollection(items);

                Assert.IsTrue(sheet.Column(1).Hidden);
                Assert.IsFalse((bool)sheet.Cells[1, 1].Value);
                Assert.IsFalse(sheet.Column(2).Hidden);
                Assert.AreEqual(1, sheet.Cells[1, 2].Value);
                Assert.IsTrue(sheet.Column(3).Hidden);
                Assert.AreEqual("Hidden 1", sheet.Cells[1, 3].Value);
                Assert.IsFalse(sheet.Column(4).Hidden);
                Assert.AreEqual("Name 1", sheet.Cells[1, 4].Value);
            }
        }

        //Testing i1416 I1416 Issue1416
        [TestMethod]
        public void NullablePropertiesShouldLoad()
        {
            using (var package = OpenPackage("LoadFromCollectionNullables.xlsx",true))
            {
                var ws = package.Workbook.Worksheets.Add("test");

                var allMembers = typeof(ColumnsWithoutAttributes).GetMembers().Where(m => m.MemberType == MemberTypes.Property).ToArray();

                ws.Cells[1, 1].LoadFromCollection(_collectionNoAttributes, PrintHeaders: true, TableStyle: TableStyles.Light1,
                    memberFlags: BindingFlags.Public | BindingFlags.Instance,
                    Members: allMembers);

                var child0 = _collectionNoAttributes[0];

                Assert.AreEqual("NullableInt", ws.Cells["A1"].Value);
                Assert.AreEqual("NonNull", ws.Cells["B1"].Value);
                Assert.AreEqual("NullableDateTime", ws.Cells["C1"].Value);
                //Nested nullable table column with property gets the property name
                Assert.AreEqual("NullableValue", ws.Cells["D1"].Value);
                Assert.AreEqual("ExplicitlyNullableString", ws.Cells["E1"].Value);
                Assert.IsNull(ws.Cells["F1"].Value);

                Assert.AreEqual(child0.NullableInt.Value, ws.Cells["A2"].Value);
                Assert.AreEqual(child0.NonNull, ws.Cells["B2"].Value);
                Assert.AreEqual(child0.NullableDateTime.Value, ws.Cells["C2"].Value);
                //Nested nullable table column with property
                Assert.AreEqual(child0.NestedNullableNullable.NullableValue.Value, ws.Cells["D2"].Value);
                Assert.AreEqual(child0.ExplicitlyNullableString, ws.Cells["E2"].Value);
                Assert.IsNull(ws.Cells["F2"].Value);

                SaveAndCleanup(package);
            }
        }
    }
}
