using EPPlusTest.RichData.TestClasses;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.RichData.IndexRelations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.RichData
{
    [TestClass]
    public class IndexedCollectionTests
    {
        [TestMethod]
        public void ReindexTest()
        {
            IdGenerator.Reset();
            var store = new RichDataIndexStore();
            var coll = new RichValueTestCollection(store);
            var item1 = new RichValueTest(store);
            var item2 = new RichValueTest(store);
            var item3 = new RichValueTest(store);
            coll.Add(item1);
            coll.Add(item2); 
            coll.Add(item3);
            var ix1 = store.GetIndexByItem(item1);
            Assert.AreEqual(0, ix1);
            var ix2 = store.GetIndexByItem(item2);
            Assert.AreEqual(1, ix2);
            var ix3 = store.GetIndexByItem(item3);
            Assert.AreEqual(2, ix3);

            // delete item2 and re-index
            item2.DeleteMe();
            store.ReIndex();
            ix1 = store.GetIndexByItem(item1);
            ix2 = store.GetIndexByItem(item2);
            ix3 = store.GetIndexByItem(item3);
            Assert.AreEqual(0, ix1);
            Assert.IsNull(ix2);
            Assert.AreEqual(1, ix3);
        }
    }
}
