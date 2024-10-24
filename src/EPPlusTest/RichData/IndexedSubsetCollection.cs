using EPPlusTest.RichData.TestClasses;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.RichData
{
    [TestClass]
    public class IndexedSubsetCollection
    {
        [TestMethod]
        public void ShouldReturnOnlyFilteredMembers()
        {
            using var package = new ExcelPackage();
            var store = new RichDataIndexStore(package.Workbook);
            var coll = new RichValueTestCollection(store);
            var item1 = new RichValueTest(store);
            var item2 = new RichValueTest(store);
            coll.Add(item1);
            coll.Add(item2);
            var subsetColl = new IndexedSubsetCollection<RichValueTest>(coll)
            {
                item1
            };
            Assert.AreEqual(2, coll.Count);
            Assert.AreEqual(1, subsetColl.Count);
            Assert.AreEqual(item1, subsetColl.First());
        }
    }
}
