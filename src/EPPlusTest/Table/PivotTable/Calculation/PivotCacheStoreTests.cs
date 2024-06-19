using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable.Calculation;
using OfficeOpenXml.Table.PivotTable.Calculation.Functions;

namespace EPPlusTest.Table.PivotTable.Calculation
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
            var store = new PivotCalculationStore();
            store.Add([1, 1], 1);
            store.Add([2, 1], 2);
            store.Add([1, 2], 3);

            Assert.AreEqual(3, store.Count);

            Assert.AreEqual(1, store[[1, 1]]);
            Assert.AreEqual(2, store[[2, 1]]);
            Assert.AreEqual(3, store[[1, 2]]);

            Assert.AreEqual(3, store.GetPreviousValue([2, 1]));
            Assert.IsNull(store.GetNextValue([2, 1]));
        }
        [TestMethod]
        public void FindPreviousNextValue()
        {
            var store = new PivotCalculationStore();
            store.Add([1, 1], 1);
            store.Add([3, 1], 2);
            store.Add([1, 3], 3);

            Assert.AreEqual(3, store.Count);

            Assert.AreEqual(1, store[[1, 1]]);
            Assert.AreEqual(2, store[[3, 1]]);
            Assert.AreEqual(3, store[[1, 3]]);

            Assert.AreEqual(3, store.GetPreviousValue([2, 1]));
            Assert.AreEqual(2, store.GetNextValue([2, 1]));
        }
        [TestMethod]
        public void FindIndex()
        {
            var store = new PivotCalculationStore();
            store.Add([3, 1, 1], 1);
            store.Add([2, 2, 2], 2);
            store.Add([2, 1, 3], 3);
            store.Add([1, 1, 1], 4);
            store.Add([1, 1, 4], 5);
            store.Add([1, 1, 2], 6);

            Assert.AreEqual(6, store.Count);

            Assert.AreEqual(1, store[[3, 1, 1]]);
            Assert.AreEqual(2, store[[2, 2, 2]]);
            Assert.AreEqual(3, store[[2, 1, 3]]);
            Assert.AreEqual(4, store[[1, 1, 1]]);
            Assert.AreEqual(5, store[[1, 1, 4]]);
            Assert.AreEqual(6, store[[1, 1, 2]]);

            Assert.AreEqual(5, store.GetIndex([3, 1, 1]));
            Assert.AreEqual(0, store.GetIndex([1, 1, 1]));

            Assert.AreEqual(-7, store.GetIndex([4, 1, 1]));
            Assert.AreEqual(-4, store.GetIndex([2, 1, 1]));
        }
		[TestMethod]
		public void PivotItemKeyTests()
		{
			//The pivot item key is used for aggregating items per row/column fields.
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, 0 }, 1));
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, 0 }, 1));
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, PivotCalculationStore.SumLevelValue }, 1));
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue }, 1));

			//2 row and 1 col
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, 0, 0 }, 2));
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, 0, PivotCalculationStore.SumLevelValue }, 2));
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, PivotCalculationStore.SumLevelValue, 0 }, 2));
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue }, 2));
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue, 0 }, 2));
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue }, 2));

			Assert.IsTrue(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, 0, PivotCalculationStore.SumLevelValue }, 2));
			Assert.IsTrue(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, 0, 0 }, 2));

			//1 row and 2 col
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, 0, 0 }, 1));
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, 0, PivotCalculationStore.SumLevelValue }, 1));
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { 0, PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue }, 1));
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, 0, PivotCalculationStore.SumLevelValue }, 1));
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, 0, 0 }, 1));
			Assert.IsFalse(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue }, 1));

			Assert.IsTrue(PivotFunction.IsNonTopLevel(new int[] { PivotCalculationStore.SumLevelValue, PivotCalculationStore.SumLevelValue, 0 }, 1));
			Assert.IsTrue(PivotFunction.IsNonTopLevel(new int[] { 0, PivotCalculationStore.SumLevelValue, 0 }, 1));
		}
	}
}
