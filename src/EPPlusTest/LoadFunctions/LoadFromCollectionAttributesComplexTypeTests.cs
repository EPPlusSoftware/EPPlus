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

namespace EPPlusTest.LoadFunctions
{
    [TestClass]
    public class LoadFromCollectionAttributesComplexTypeTests
    {
        private List<Outer> _collection = new List<Outer>();
        private List<OuterWithHeaders> _collectionHeaders = new List<OuterWithHeaders>();
        private List<OuterReversedSortOrder> _collectionReversed = new List<OuterReversedSortOrder>();
        private List<OuterSubclass> _collectionInheritence = new List<OuterSubclass>();

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
        }

        [TestCleanup]
        public void Cleanup()
        {
            _collection.Clear();
        }

        [TestMethod]
        public void ShouldSetupColumnsWithPath()
        {
            var cols = new LoadFromCollectionColumns<Outer>(LoadFromCollectionParams.DefaultBindingFlags, Enumerable.Empty<string>().ToList());
            var result = cols.Setup();
            Assert.AreEqual(5, result.Count, "List did not contain 5 elements as expected");
            Assert.AreEqual("ApprovedUtc", result[0].Path);
            Assert.AreEqual("Organization.OrgLevel3", result[1].Path);
        }

        [TestMethod]
        public void ShouldSetupColumnsWithPathSorted()
        {
            var cols = new LoadFromCollectionColumns<OuterReversedSortOrder>(LoadFromCollectionParams.DefaultBindingFlags);
            var result = cols.Setup();
            Assert.AreEqual(5, result.Count, "List did not contain 5 elements as expected");
            Assert.AreEqual("Acknowledged", result[0].Path);
            Assert.AreEqual("Organization.OrgLevel5", result[1].Path);
            Assert.AreEqual("ApprovedUtc", result.Last().Path);
        }

        [TestMethod]
        public void ShouldSetupColumnsWithPathSortedByClassAttribute()
        {
            var order = new List<string>
            {
                "ApprovedUtc",
                "Acknowledged",
                "Organization.OrgLevel5"
            };
            var cols = new LoadFromCollectionColumns<OuterReversedSortOrder>(LoadFromCollectionParams.DefaultBindingFlags, order);
            var result = cols.Setup();
            Assert.AreEqual(5, result.Count, "List did not contain 5 elements as expected");
            Assert.AreEqual("ApprovedUtc", result[0].Path);
            Assert.AreEqual("Acknowledged", result[1].Path);
            Assert.AreEqual("Organization.OrgLevel5", result[2].Path);
            Assert.AreEqual("Organization.OrgLevel4", result[3].Path);
            Assert.AreEqual("Organization.OrgLevel3", result[4].Path);

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
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("test");
                ws.Cells["A1"].LoadFromCollection(_collection);
                Assert.IsNull(ws.Cells["B1"].Value);
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
                var cv = ws.Cells["F1"].Value;
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
                var cv = ws.Cells["G1"].Value;
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
            using(var package = new ExcelPackage())
            {
                var items = ExcelItems.GetItems1();
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].LoadFromCollection(items);
                Assert.AreEqual("Product Family", sheet.Cells["A1"].Value);
                Assert.AreEqual("PCH Die Name", sheet.Cells["B1"].Value);
                Assert.AreEqual("Collateral Owner Email", sheet.Cells["F1"].Value);
                Assert.AreEqual("Mission Control Lead Email", sheet.Cells["I1"].Value);
                Assert.AreEqual("Created (GMT)", sheet.Cells["L1"].Value);
            }
        }
    }
}
