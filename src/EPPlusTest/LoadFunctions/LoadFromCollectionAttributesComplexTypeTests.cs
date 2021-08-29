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
        [EpplusTable]
        public class Organization
        {
            [EpplusTableColumn(Header = "Org Level 3", Order = 1)]
            public string OrgLevel3 { get; set; }

            [EpplusTableColumn(Header = "Org Level 4", Order = 2)]
            public string OrgLevel4 { get; set; }

            [EpplusTableColumn(Header = "Org Level 5", Order = 3)]
            public string OrgLevel5 { get; set; }
        }

        [EpplusTable]
        public class OrganizationReversedSortOrder
        {
            [EpplusTableColumn(Header = "Org Level 3", Order = 3)]
            public string OrgLevel3 { get; set; }

            [EpplusTableColumn(Header = "Org Level 4", Order = 2)]
            public string OrgLevel4 { get; set; }

            [EpplusTableColumn(Header = "Org Level 5", Order = 1)]
            public string OrgLevel5 { get; set; }
        }

        [EpplusTable]
        public class Outer
        {
            [EpplusTableColumn(Header = nameof(ApprovedUtc), Order = 1)]
            public DateTime? ApprovedUtc { get; set; }

            [EpplusNestedTableColumn(Order = 2)]
            public Organization Organization { get; set; }

            [EpplusTableColumn(Header = "Acknowledged...", Order = 3)]
            public bool Acknowledged { get; set; }
        }

        [EpplusTable]
        public class OuterReversedSortOrder
        {
            [EpplusTableColumn(Header = nameof(ApprovedUtc), Order = 3)]
            public DateTime? ApprovedUtc { get; set; }

            [EpplusNestedTableColumn(Order = 2)]
            public OrganizationReversedSortOrder Organization { get; set; }

            [EpplusTableColumn(Header = "Acknowledged...", Order = 1)]
            public bool Acknowledged { get; set; }
        }

        private List<Outer> _collection = new List<Outer>();
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
    }
}
