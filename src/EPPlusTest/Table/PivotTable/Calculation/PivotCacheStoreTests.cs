using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable.Calculation;

namespace EPPlusTest.Table.PivotTable
{
    [TestClass]
    public class PivotCacheStoreTests : TestBase
    {
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
        }
        [ClassCleanup]
        public static void Cleanup()
        {
        }
        [TestMethod]
        public void FindExactValue()
        {
            var store = new PivotCacheStore();
            store.Add([1, 1], 1);
            store.Add([2, 1], 2);
            store.Add([1, 2], 3);

            Assert.AreEqual(3, store.Count);

            Assert.AreEqual(1, store[[1, 1]]);
            Assert.AreEqual(2, store[[2, 1]]);
            Assert.AreEqual(3, store[[1, 2]]);

            Assert.AreEqual(1, store.GetPreviousValue([2, 1]));
            Assert.AreEqual(3, store.GetNextValue([2, 1]));
        }
    }
}
