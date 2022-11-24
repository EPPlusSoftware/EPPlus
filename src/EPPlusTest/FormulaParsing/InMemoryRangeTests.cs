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
            var address = new FormulaRangeAddress(_context) { FromCol = 1, ToCol = 1, FromRow = 1, ToRow = 1, WorksheetIx = 0 };
            _context.Scopes.NewScope(address);
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
        public void GetNCellsSholdReturnCorrectResult2_2()
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
        public void GetNCellsSholdReturnCorrectResult2_1()
        {
            var rangeDef = new RangeDefinition(3, 1);
            var range = new InMemoryRange(rangeDef);
            range.SetValue(0, 0, 1);
            range.SetValue(1, 0, 2);
            range.SetValue(2, 0, 3);
            var nCells = range.GetNCells();
            Assert.AreEqual(3, nCells);
        }
    }
}
