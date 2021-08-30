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
        }

        [TestCleanup]
        public void Cleanup()
        {
            _collection.Clear();
        }

        [TestMethod]
        public void ShouldSetupColumnsWithPath()
        {
            var cols = new LoadFromCollectionColumns<Outer>(LoadFromCollectionParams.DefaultBindingFlags);
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
    }
}
