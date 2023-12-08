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
            store.Add(new int[] { 1, 1 }, 1);
            store.Add(new int[] { 2, 1 }, 2);
            store.Add(new int[] { 1, 2 }, 3);

            Assert.AreEqual(3, store.Count);
        }
    }
}
