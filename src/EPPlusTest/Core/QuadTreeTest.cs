using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Diagnostics;

namespace EPPlusTest.Core
{
    [TestClass]
    public class QuadTreeTest
    {
        [TestMethod]
        public void QuadTreeIntersect1Test()
        {
            var qt = new QuadTree<int>(1,1,5000,200);
            qt.Add(new QuadRange(2, 1, 50, 20), 1);
            qt.Add(new QuadRange(55, 7, 55, 7), 2);

            qt.Add(new QuadRange(2400, 8, 2900, 12), 3);


            var range1 = new QuadRange(44, 2, 88, 100);
            var ranges = qt.GetIntersectingRanges(range1);

            Assert.AreEqual(2, ranges.Count);
        }
        [TestMethod]
        public void QuadTreeIntersectAboveAndBelowTest()
        {
            var qt = new QuadTree<int>(1, 1, 5000, 500);

            qt.Add(new QuadRange(1000, 100, 1000, 100), 1);
            qt.Add(new QuadRange(1010, 95, 1020, 105), 2);
            qt.Add(new QuadRange(900, 80, 1100, 120), 3);
            qt.Add(new QuadRange(1100, 95, 1020, 105), 4);
            qt.Add(new QuadRange(500, 50, 2000, 200), 5);
            //Not intersecting
            qt.Add(new QuadRange(1, 20, 899, 20), 6);
            qt.Add(new QuadRange(1101, 20, 1200, 20), 7);

            var ranges = qt.GetIntersectingRanges(new QuadRange(900, 50, 1100, 108));
            Assert.AreEqual(5, ranges.Count);
        }
        [TestMethod]
        public void QuadLargeTest()
        {
            var rows = 1000000;
            var cols = 1000;
            var qt = new QuadTree<int>(1, 1, rows, cols);
            var sw = new Stopwatch();
            sw.Start();
            var items = AddRangeItems(rows, cols, qt, 50, 50);
            sw.Stop();
            Debug.WriteLine($"Added {items} items  in {sw.ElapsedMilliseconds} ms");
            sw.Restart();
            
            var r1 = new QuadRange(5000, 200, 10000, 300);
            var ir1 = qt.GetIntersectingRangeItems(r1);
            foreach(var r in ir1)
            {
                if(r.Range.Intersect(r1)==IntersectType.OutSide)
                {
                    Assert.Fail($"Range {r.Range} does not intersect with {r1}");
                }
            }
            sw.Stop();
            Debug.WriteLine($"Queried {ir1.Count} items in {sw.ElapsedMilliseconds} ms");

        }
        private static int AddRangeItems(int rows, int cols, QuadTree<int> qt, int rowsIntervall, int colIntervall)
        {
            var count = 0;
            for (int r = 1; r < rows; r += rowsIntervall)
            {
                for (int c = 1; c < cols; c += colIntervall)
                {
                    qt.Add(new QuadRange(r, c, r + rowsIntervall, c + colIntervall), r * c + c);
                    count++;
                }
            }
            return count;
        }
    }
}
