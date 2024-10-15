using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
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
        public void QuadTreeAddressTest()
        {
            var qt = new QuadTree<int>(new ExcelAddress("A1:C20"));

            qt.Add(new QuadRange(new ExcelAddress("A3")), 1);
            qt.Add(new QuadRange(new ExcelAddress("B5:B10")), 2);
            qt.Add(new QuadRange(new ExcelAddress("C2:C3")), 3);

            //Not intersecting
            qt.Add(new QuadRange(new ExcelAddress("D3")), 4);
            qt.Add(new QuadRange(new ExcelAddress("Z5:Z10")), 5);
            qt.Add(new QuadRange(new ExcelAddress("F2:F3")), 6);

            var ranges = qt.GetIntersectingRanges(new QuadRange(new ExcelAddress("B3:E40")));
            Assert.AreEqual(3, ranges.Count);
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
        [TestMethod]
        public void QuadTree_InsertRows_Inside()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(5, 5, 10, 10);
            qt.Add(r, 1);
            qt.InsertRow(5, 2);
            var ranges = qt.GetIntersectingRanges(new QuadRange(13, 5, 11, 5));
        }
        [TestMethod]
        public void QuadTree_InsertRows_Expand()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(30, 1, 35, 5);
            qt.Add(r, 1);
            qt.InsertRow(5, 2);
            var ranges = qt.GetIntersectingRanges(new QuadRange(30, 5, 32, 5));
            Assert.AreEqual(1, ranges.Count);
        }
        [TestMethod]
        public void QuadTree_InsertRows_ExpandMulti()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(500, 1, 505, 5);
            qt.Add(r, 1);
            qt.InsertRow(5, 2);
            var ranges = qt.GetIntersectingRanges(new QuadRange(500, 5, 505, 5));
            Assert.AreEqual(1, ranges.Count);
        }
        [TestMethod]
        public void QuadTree_DeleteRow()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(10, 10, 12, 10);
            qt.Add(r, 1);
            qt.DeleteRow(2, 2);
            var ranges = qt.GetIntersectingRanges(new QuadRange(8, 10, 10, 10));
            Assert.AreEqual(1, ranges.Count);
            Assert.AreEqual(8, ranges[0].FromRow);
            Assert.AreEqual(10, ranges[0].ToRow);

            qt.DeleteRow(6, 3);
            ranges = qt.GetIntersectingRanges(new QuadRange(6, 10, 6, 10));

            Assert.AreEqual(1, ranges.Count);
            Assert.AreEqual(6, ranges[0].FromRow);
            Assert.AreEqual(7, ranges[0].ToRow);
        }
        [TestMethod]
        public void QuadTree_DeleteRowMoveToQuad()
        {
            var qt = new QuadTree<int>(1,1,60,60);
            var r = new QuadRange(29, 29, 31, 29);
            qt.Add(r, 1);
            Assert.AreEqual(1, qt.Root.Ranges.Count);
            qt.DeleteRow(2, 2);
            Assert.AreEqual(0, qt.Root.Ranges.Count);
            Assert.AreEqual(1, qt.Root.Quads[0].Ranges.Count);

            var ranges = qt.GetIntersectingRanges(new QuadRange(27, 29, 29, 29));
        }
        [TestMethod]
        public void QuadTree_ClearTopLeft()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(5, 5, 10, 10);
            qt.Add(r, 1);

            qt.Clear(4, 4, 7, 7);
            Assert.AreEqual(2, qt.Root.Ranges.Count);
            Assert.AreEqual("H5:J10", qt.Root.Ranges[0].Range.ToString());
            Assert.AreEqual("E8:G10", qt.Root.Ranges[1].Range.ToString());
        }
        [TestMethod]
        public void QuadTree_ClearTop()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(5, 5, 10, 10);
            qt.Add(r, 1);

            qt.Clear(4, 5, 6, 10);
            Assert.AreEqual(1, qt.Root.Ranges.Count);
            Assert.AreEqual("E7:J10", qt.Root.Ranges[0].Range.ToString());
        }
        [TestMethod]
        public void QuadTree_ClearTopRight()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(5, 5, 10, 10);
            qt.Add(r, 1);

            qt.Clear(4, 8, 8, 12);
            Assert.AreEqual(2, qt.Root.Ranges.Count);
            Assert.AreEqual("E5:G10", qt.Root.Ranges[0].Range.ToString());
            Assert.AreEqual("H9:J10", qt.Root.Ranges[1].Range.ToString());
        }
        [TestMethod]
        public void QuadTree_ClearLeft()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(5, 5, 10, 10);
            qt.Add(r, 1);

            qt.Clear(4, 5, 13, 7);
            Assert.AreEqual(1, qt.Root.Ranges.Count);
            Assert.AreEqual("H5:J10", qt.Root.Ranges[0].Range.ToString());
        }
        [TestMethod]
        public void QuadTree_ClearRight()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(5, 5, 10, 10);
            qt.Add(r, 1);

            qt.Clear(5, 9, 13, 12);
            Assert.AreEqual(1, qt.Root.Ranges.Count);
            Assert.AreEqual("E5:H10", qt.Root.Ranges[0].Range.ToString());
        }
        [TestMethod]
        public void QuadTree_ClearBottomRight()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(5, 5, 10, 10);
            qt.Add(r, 1);

            qt.Clear(8, 4, 13, 7);
            Assert.AreEqual(2, qt.Root.Ranges.Count);
            Assert.AreEqual("E5:J7", qt.Root.Ranges[0].Range.ToString());
            Assert.AreEqual("H8:J10", qt.Root.Ranges[1].Range.ToString());
        }
        [TestMethod]
        public void QuadTree_ClearBottom()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(5, 5, 10, 10);
            qt.Add(r, 1);

            qt.Clear(10, 5, 13, 10);
            Assert.AreEqual(1, qt.Root.Ranges.Count);
            Assert.AreEqual("E5:J9", qt.Root.Ranges[0].Range.ToString());
        }
        [TestMethod]
        public void QuadTree_ClearBottomLeft()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(5, 5, 10, 10);
            qt.Add(r, 1);

            qt.Clear(8, 8, 13, 15);
            Assert.AreEqual(2, qt.Root.Ranges.Count);
            Assert.AreEqual("E5:J7", qt.Root.Ranges[0].Range.ToString());
            Assert.AreEqual("E8:G10", qt.Root.Ranges[1].Range.ToString());
        }
        [TestMethod]
        public void QuadTree_ClearHorizontal()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(5, 5, 10, 10);
            qt.Add(r, 1);

            qt.Clear(7, 5, 8, 10);
            Assert.AreEqual(2, qt.Root.Ranges.Count);
            Assert.AreEqual("E5:J6", qt.Root.Ranges[0].Range.ToString());
            Assert.AreEqual("E9:J10", qt.Root.Ranges[1].Range.ToString());
        }
        [TestMethod]
        public void QuadTree_ClearVertical()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(5, 5, 10, 10);
            qt.Add(r, 1);

            qt.Clear(5, 7, 10, 8);
            Assert.AreEqual(2, qt.Root.Ranges.Count);
            Assert.AreEqual("E5:F10", qt.Root.Ranges[0].Range.ToString());
            Assert.AreEqual("I5:J10", qt.Root.Ranges[1].Range.ToString());
        }

        [TestMethod]
        public void QuadTree_ClearInside()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(5, 5, 10, 10);
            qt.Add(r, 1);

            qt.Clear(8, 8, 8, 8);
            Assert.AreEqual(4, qt.Root.Ranges.Count);
            Assert.AreEqual("E5:J7", qt.Root.Ranges[0].Range.ToString());
            Assert.AreEqual("E8:G10", qt.Root.Ranges[1].Range.ToString());
            Assert.AreEqual("I8:J10", qt.Root.Ranges[2].Range.ToString());
            Assert.AreEqual("H9:H10", qt.Root.Ranges[3].Range.ToString());
        }
        [TestMethod]
        public void QuadTree_ClearAll()
        {
            var qt = new QuadTree<int>();
            var r = new QuadRange(5, 5, 10, 10);
            qt.Add(r, 1);

            qt.Clear(5, 5, 10, 10);
            Assert.AreEqual(0, qt.Root.Ranges.Count);
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
