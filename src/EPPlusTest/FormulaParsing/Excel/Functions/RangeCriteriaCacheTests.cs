using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Collections.Generic;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestClass]
    public class RangeCriteriaCacheTests
    {
        private ExcelPackage _package;
        private RangeCriteriaCache _cache;

        [TestInitialize]
        public void Setup()
        {
            _package = new ExcelPackage();
            _package.Workbook.Worksheets.Add("Sheet1");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package?.Dispose();
        }

        [TestMethod]
        public void FlattenedRange_ShouldCacheAndRetrieve()
        {
            _cache = new RangeCriteriaCache(_package);
            var address = new FormulaRangeAddress { WorksheetIx = 0, FromRow = 1, FromCol = 1, ToRow = 10, ToCol = 1 };
            var data = new List<object> { 1, 2, 3, 4, 5 };

            _cache.SetFlattenedRange(address, data);
            var retrieved = _cache.GetFlattenedRange(address);

            Assert.IsNotNull(retrieved);
            Assert.AreEqual(5, retrieved.Count);
            Assert.AreEqual(1, retrieved[0]);
        }

        [TestMethod]
        public void FlattenedRange_FIFO_ShouldEvictOldest()
        {
            _cache = new RangeCriteriaCache(_package, maxFlattenedRanges: 3);

            var address1 = new FormulaRangeAddress { WorksheetIx = 0, FromRow = 1, FromCol = 1, ToRow = 10, ToCol = 1 };
            var address2 = new FormulaRangeAddress { WorksheetIx = 0, FromRow = 11, FromCol = 1, ToRow = 20, ToCol = 1 };
            var address3 = new FormulaRangeAddress { WorksheetIx = 0, FromRow = 21, FromCol = 1, ToRow = 30, ToCol = 1 };
            var address4 = new FormulaRangeAddress { WorksheetIx = 0, FromRow = 31, FromCol = 1, ToRow = 40, ToCol = 1 };

            var data1 = new List<object> { 1 };
            var data2 = new List<object> { 2 };
            var data3 = new List<object> { 3 };
            var data4 = new List<object> { 4 };

            _cache.SetFlattenedRange(address1, data1);
            _cache.SetFlattenedRange(address2, data2);
            _cache.SetFlattenedRange(address3, data3);

            // Cache is now full (3/3)
            Assert.IsNotNull(_cache.GetFlattenedRange(address1), "Address1 should be in cache");
            Assert.IsNotNull(_cache.GetFlattenedRange(address2), "Address2 should be in cache");
            Assert.IsNotNull(_cache.GetFlattenedRange(address3), "Address3 should be in cache");

            // Adding 4th item should evict the oldest (address1)
            _cache.SetFlattenedRange(address4, data4);

            Assert.IsNull(_cache.GetFlattenedRange(address1), "Address1 should be evicted (oldest)");
            Assert.IsNotNull(_cache.GetFlattenedRange(address2), "Address2 should still be in cache");
            Assert.IsNotNull(_cache.GetFlattenedRange(address3), "Address3 should still be in cache");
            Assert.IsNotNull(_cache.GetFlattenedRange(address4), "Address4 should be in cache");
        }

        [TestMethod]
        public void MatchIndexes_ShouldCacheAndRetrieve()
        {
            _cache = new RangeCriteriaCache(_package);
            var address = new FormulaRangeAddress { WorksheetIx = 0, FromRow = 1, FromCol = 1, ToRow = 10, ToCol = 1 };
            var criteria = "test";
            var indexes = new List<int> { 1, 3, 5, 7 };

            _cache.SetMatchIndexes(address, criteria, indexes);
            var retrieved = _cache.GetMatchIndexes(address, criteria);

            Assert.IsNotNull(retrieved);
            Assert.AreEqual(4, retrieved.Count);
            Assert.AreEqual(1, retrieved[0]);
            Assert.AreEqual(7, retrieved[3]);
        }

        [TestMethod]
        public void MatchIndexes_FIFO_ShouldEvictOldest()
        {
            _cache = new RangeCriteriaCache(_package, maxMatchIndexes: 3);

            var address = new FormulaRangeAddress { WorksheetIx = 0, FromRow = 1, FromCol = 1, ToRow = 10, ToCol = 1 };
            var criteria1 = "A";
            var criteria2 = "B";
            var criteria3 = "C";
            var criteria4 = "D";

            var indexes1 = new List<int> { 1 };
            var indexes2 = new List<int> { 2 };
            var indexes3 = new List<int> { 3 };
            var indexes4 = new List<int> { 4 };

            _cache.SetMatchIndexes(address, criteria1, indexes1);
            _cache.SetMatchIndexes(address, criteria2, indexes2);
            _cache.SetMatchIndexes(address, criteria3, indexes3);

            // Cache is now full (3/3)
            Assert.IsNotNull(_cache.GetMatchIndexes(address, criteria1), "Criteria1 should be in cache");
            Assert.IsNotNull(_cache.GetMatchIndexes(address, criteria2), "Criteria2 should be in cache");
            Assert.IsNotNull(_cache.GetMatchIndexes(address, criteria3), "Criteria3 should be in cache");

            // Adding 4th item should evict the oldest (criteria1)
            _cache.SetMatchIndexes(address, criteria4, indexes4);

            Assert.IsNull(_cache.GetMatchIndexes(address, criteria1), "Criteria1 should be evicted (oldest)");
            Assert.IsNotNull(_cache.GetMatchIndexes(address, criteria2), "Criteria2 should still be in cache");
            Assert.IsNotNull(_cache.GetMatchIndexes(address, criteria3), "Criteria3 should still be in cache");
            Assert.IsNotNull(_cache.GetMatchIndexes(address, criteria4), "Criteria4 should be in cache");
        }

        [TestMethod]
        public void MatchIndexes_WithFormulas_ShouldNotCache()
        {
            var ws = _package.Workbook.Worksheets[0];
            ws.Cells["A1"].Formula = "=1+1";
            ws.Cells["A2"].Value = 2;

            _cache = new RangeCriteriaCache(_package);
            var address = new FormulaRangeAddress { WorksheetIx = 0, FromRow = 1, FromCol = 1, ToRow = 2, ToCol = 1 };
            var criteria = "test";
            var indexes = new List<int> { 1 };

            _cache.SetMatchIndexes(address, criteria, indexes);

            // Should not cache because range has formulas
            var retrieved = _cache.GetMatchIndexes(address, criteria);
            Assert.IsNull(retrieved, "Should not cache ranges with formulas");
        }

        [TestMethod]
        public void MatchIndexes_DifferentCriteria_SamRange_ShouldCacheSeparately()
        {
            _cache = new RangeCriteriaCache(_package);
            var address = new FormulaRangeAddress { WorksheetIx = 0, FromRow = 1, FromCol = 1, ToRow = 10, ToCol = 1 };

            var indexes1 = new List<int> { 1, 2 };
            var indexes2 = new List<int> { 3, 4 };

            _cache.SetMatchIndexes(address, "criteriaA", indexes1);
            _cache.SetMatchIndexes(address, "criteriaB", indexes2);

            var retrieved1 = _cache.GetMatchIndexes(address, "criteriaA");
            var retrieved2 = _cache.GetMatchIndexes(address, "criteriaB");

            Assert.IsNotNull(retrieved1);
            Assert.IsNotNull(retrieved2);
            Assert.AreEqual(2, retrieved1.Count);
            Assert.AreEqual(2, retrieved2.Count);
            Assert.AreEqual(1, retrieved1[0]);
            Assert.AreEqual(3, retrieved2[0]);
        }

        [TestMethod]
        public void Clear_ShouldRemoveAllCachedData()
        {
            _cache = new RangeCriteriaCache(_package);
            var address = new FormulaRangeAddress { WorksheetIx = 0, FromRow = 1, FromCol = 1, ToRow = 10, ToCol = 1 };

            _cache.SetFlattenedRange(address, new List<object> { 1, 2, 3 });
            _cache.SetMatchIndexes(address, "test", new List<int> { 1, 2 });

            Assert.IsNotNull(_cache.GetFlattenedRange(address));
            Assert.IsNotNull(_cache.GetMatchIndexes(address, "test"));

            _cache.Clear();

            Assert.IsNull(_cache.GetFlattenedRange(address));
            Assert.IsNull(_cache.GetMatchIndexes(address, "test"));
        }

        [TestMethod]
        public void FlattenedRange_UpdateExisting_ShouldNotDuplicate()
        {
            _cache = new RangeCriteriaCache(_package, maxFlattenedRanges: 2);
            var address = new FormulaRangeAddress { WorksheetIx = 0, FromRow = 1, FromCol = 1, ToRow = 10, ToCol = 1 };

            var data1 = new List<object> { 1, 2, 3 };
            var data2 = new List<object> { 4, 5, 6 };

            // Add same address twice with different data
            _cache.SetFlattenedRange(address, data1);
            _cache.SetFlattenedRange(address, data2);

            var retrieved = _cache.GetFlattenedRange(address);

            // Should have updated value, not duplicated
            Assert.IsNotNull(retrieved);
            Assert.AreEqual(4, retrieved[0], "Should have updated data");
        }

        [TestMethod]
        public void MatchIndexes_UpdateExisting_ShouldNotDuplicate()
        {
            _cache = new RangeCriteriaCache(_package, maxMatchIndexes: 2);
            var address = new FormulaRangeAddress { WorksheetIx = 0, FromRow = 1, FromCol = 1, ToRow = 10, ToCol = 1 };
            var criteria = "test";

            var indexes1 = new List<int> { 1, 2 };
            var indexes2 = new List<int> { 3, 4 };

            // Add same address+criteria twice with different data
            _cache.SetMatchIndexes(address, criteria, indexes1);
            _cache.SetMatchIndexes(address, criteria, indexes2);

            var retrieved = _cache.GetMatchIndexes(address, criteria);

            // Should have updated value, not duplicated
            Assert.IsNotNull(retrieved);
            Assert.AreEqual(3, retrieved[0], "Should have updated indexes");
        }
    }
}