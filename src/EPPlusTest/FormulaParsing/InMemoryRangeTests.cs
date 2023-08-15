using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class InMemoryRangeTests
    {
        private ParsingContext _context;
        private ExcelPackage _package;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _package.Workbook.Worksheets.Add("test");
            _context = ParsingContext.Create(_package);
            _context.CurrentCell = new FormulaCellAddress(0, 1, 1);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void ShouldCreateSimpleMatrix()
        {
            var rangeDef = new RangeDefinition(2, 2);
            var range = new InMemoryRange(rangeDef);
            range.SetValue(0, 0, 1);
            range.SetValue(0, 1, 2);
            range.SetValue(1, 0, 3);
            var v1 = range.GetValue(0, 0);
            var v2 = range.GetValue(0, 1);
            var v3 = range.GetValue(1, 0);
            var v4 = range.GetValue(1, 1);
            Assert.AreEqual(1, v1);
            Assert.AreEqual(2, v2);
            Assert.AreEqual(3, v3);
            Assert.IsNull(v4);
        }

        [TestMethod]
        public void EnumerableOrderShouldBeCorrect()
        {
            var rangeDef = new RangeDefinition(2, 2);
            var range = new InMemoryRange(rangeDef);
            range.SetValue(0, 0, 1);
            range.SetValue(1, 0, 2);
            range.SetValue(0, 1, 3);
            var lst = range.ToList();
            Assert.AreEqual(4, lst.Count);
            var v1 = lst[0].Value;
            var v2 = lst[1].Value;
            var v3 = lst[2].Value;
            var v4 = lst[3];
            Assert.AreEqual(1, v1);
            Assert.AreEqual(3, v2);
            Assert.AreEqual(2, v3);
            Assert.IsNull(v4);
        }

        [TestMethod]
        public void GetNCellsShouldReturnCorrectResult2_2()
        {
            var rangeDef = new RangeDefinition(2, 2);
            var range = new InMemoryRange(rangeDef);
            range.SetValue(0, 0, 1);
            range.SetValue(1, 0, 2);
            range.SetValue(0, 1, 3);
            var nCells = range.GetNCells();
            Assert.AreEqual(4, nCells);
        }

        [TestMethod]
        public void GetNCellsShouldReturnCorrectResult2_1()
        {
            var rangeDef = new RangeDefinition(3, 1);
            var range = new InMemoryRange(rangeDef);
            range.SetValue(0, 0, 1);
            range.SetValue(1, 0, 2);
            range.SetValue(2, 0, 3);
            var nCells = range.GetNCells();
            Assert.AreEqual(3, nCells);
        }

        [TestMethod]
        public void GetOffsetMultiRangeShouldReturnCorrectResult()
        {
            var rangeDef = new RangeDefinition(3, 3);
            var range = new InMemoryRange(rangeDef);
            range.SetValue(0, 0, 1);
            range.SetValue(1, 0, 2);
            range.SetValue(2, 0, 3);
            
            range.SetValue(0, 1, 4);
            range.SetValue(1, 1, 5);
            range.SetValue(2, 1, 6);
            
            range.SetValue(0, 2, 7);
            range.SetValue(1, 2, 8);
            range.SetValue(2, 2, 9);

            var newRange = range.GetOffset(1, 1, 2, 2);
            var topLeft = newRange.GetOffset(0, 0);
            var bottomRight = newRange.GetOffset(1, 1);
            Assert.AreEqual(5, topLeft);
            Assert.AreEqual(9, bottomRight);
            Assert.AreEqual(2, newRange.Size.NumberOfRows);
            Assert.AreEqual(2, newRange.Size.NumberOfCols);
        }

        [TestMethod]
        public void GetOffsetSingleRangeShouldReturnCorrectResult()
        {
            var rangeDef = new RangeDefinition(3, 3);
            var range = new InMemoryRange(rangeDef);
            range.SetValue(0, 0, 1);

            var newRange = range.GetOffset(0, 0, 0, 0);
            var topLeft = newRange.GetOffset(0, 0);
            Assert.AreEqual(1, topLeft);
            Assert.AreEqual(1, newRange.Size.NumberOfRows);
            Assert.AreEqual(1, newRange.Size.NumberOfCols);
        }

        [TestMethod]
        public void GetOffsetSingleColShouldReturnCorrectResult()
        {
            var rangeDef = new RangeDefinition(3, 3);
            var range = new InMemoryRange(rangeDef);
            range.SetValue(0, 0, 1);
            range.SetValue(1, 0, 2);

            var newRange = range.GetOffset(0, 0, 1, 0);
            var top = newRange.GetOffset(0, 0);
            Assert.AreEqual(1, top);
            var bottom = newRange.GetOffset(1, 0);
            Assert.AreEqual(2, bottom);
            Assert.AreEqual(2, newRange.Size.NumberOfRows);
            Assert.AreEqual(1, newRange.Size.NumberOfCols);
        }

        [TestMethod]
        public void GetOffsetSingleRowShouldReturnCorrectResult()
        {
            var rangeDef = new RangeDefinition(3, 3);
            var range = new InMemoryRange(rangeDef);
            range.SetValue(0, 0, 1);
            range.SetValue(0, 1, 2);

            var newRange = range.GetOffset(0, 0, 0, 1);
            var left = newRange.GetOffset(0, 0);
            Assert.AreEqual(1, left);
            var right = newRange.GetOffset(0, 1);
            Assert.AreEqual(2, right);
            Assert.AreEqual(1, newRange.Size.NumberOfRows);
            Assert.AreEqual(2, newRange.Size.NumberOfCols);
        }

        [TestMethod]
        public void TransposeShouldReturnTransposedRange()
        {
            var range = new InMemoryRange(3, 2);
            range.SetValue(0, 0, 1);
            range.SetValue(1, 0, 2);
            range.SetValue(2, 0, 3); 
            range.SetValue(0, 1, 4);
            range.SetValue(1, 1, 5);
            range.SetValue(2, 1, 6);

            var tr = range.Transpose();

            Assert.AreEqual(2, tr.Size.NumberOfRows);
            Assert.AreEqual(3, tr.Size.NumberOfCols);

            Assert.AreEqual(1, tr.GetValue(0, 0));
            Assert.AreEqual(4, tr.GetValue(1, 0));
            Assert.AreEqual(2, tr.GetValue(0, 1));
            Assert.AreEqual(5, tr.GetValue(1, 1));
            Assert.AreEqual(3, tr.GetValue(0, 2));
            Assert.AreEqual(6, tr.GetValue(1, 2));


        }
    }
}
