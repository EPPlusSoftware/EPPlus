using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core
{
    [TestClass]
    public class RangeDictionaryTests : TestBase
    {
        [ClassInitialize]
        public static void Init(TestContext context)
        {
        }
        private static int GetFromRow(long address)
        {
            return (int)(address >> 20) + 1;
        }
        private static int GetToRow(long address)
        {
            return (int)(address & 0xFFFFF) + 1; 
        }

        [TestMethod]
        public void VerifyAddAddress1()
        {
            var rd = new RangeDictionary<int>();

            rd.Add(1,1,5,5, 1);

            Assert.IsTrue(rd.Exists(1, 1));
            Assert.IsTrue(rd.Exists(2, 2));
            Assert.IsTrue(rd.Exists(5, 5));
            Assert.IsFalse(rd.Exists(6, 5));
            Assert.IsFalse(rd.Exists(5, 6));

            Assert.AreEqual(1, rd[1, 1]);
            Assert.AreEqual(1, rd[3, 3]);
            Assert.AreEqual(1, rd[5, 5]);
        }
        [TestMethod]
        public void VerifyAddAddressFillGap()
        {
            var rd = new RangeDictionary<int>();

            rd.Add(1,1,5,5, 1);
            rd.Add(6,1,7,5, 2);
            rd.Add(8,1,15,5, 3);

            Assert.IsTrue(rd.Exists(1, 1));
            Assert.IsTrue(rd.Exists(2, 2));
            Assert.IsTrue(rd.Exists(5, 5));
            Assert.IsTrue(rd.Exists(6, 4));
            Assert.IsTrue(rd.Exists(8, 3));

            Assert.AreEqual(1, rd[1, 1]);
            Assert.AreEqual(1, rd[5, 2]);
            Assert.AreEqual(2, rd[6, 3]);
            Assert.AreEqual(2, rd[7, 4]);
            Assert.AreEqual(3, rd[8, 5]);
        }
        [TestMethod]
        public void VerifyAddAddressWithSpan()
        {
            var rd = new RangeDictionary<int>();

            rd.Add(1,1,2,5, 1);
            //var r2 = new FormulaRangeAddress() { FromRow = 6, ToRow = 7, FromCol = 1, ToCol = 5 };
            rd.Add(6,1,7,5, 2);
            //var r3 = new FormulaRangeAddress() { FromRow = 9, ToRow = 15, FromCol = 1, ToCol = 5 };
            rd.Add(9, 1, 15, 5, 3);

            //var r4 = new FormulaRangeAddress() { FromRow = 4, ToRow = 4, FromCol = 1, ToCol = 5 };
            rd.Add(4,1,4,5, 4);

            //var r5 = new FormulaRangeAddress() { FromRow = 8, ToRow = 8, FromCol = 1, ToCol = 5 };
            rd.Add(8, 1,8, 5, 5);

            Assert.IsTrue(rd.Exists(1, 1));
            Assert.IsTrue(rd.Exists(2, 2));
            Assert.IsFalse(rd.Exists(5, 5));
            Assert.IsTrue(rd.Exists(6, 4));
            Assert.IsTrue(rd.Exists(9, 3));
            Assert.IsTrue(rd.Exists(15, 3));
            Assert.IsTrue(rd.Exists(8, 3));

            Assert.AreEqual(1, rd[1, 1]);
            Assert.AreEqual(4, rd[4, 2]);
            Assert.AreEqual(0, rd[5, 2]);
            Assert.AreEqual(2, rd[6, 3]);
            Assert.AreEqual(5, rd[8, 4]);
            Assert.AreEqual(3, rd[12, 5]);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void VerifyOverlapBottomRightThrowsException()
        {
            var rd = new RangeDictionary<int>();

            rd.Add(1,1,5,5, 1);
            rd.Add(5,5,6,6, 2);
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void VerifyOverlapTopLeftThrowsException()
        {
            var rd = new RangeDictionary<int>();

            rd.Add(5,5,6,6, 2);
            rd.Add(1,1,5,5, 1);
        }
        [TestMethod]
        public void VerifyInsertOnRow()
        {
            var rd = new RangeDictionary<int>();

            rd.Add(1, 1, 5, 5, 1);
            rd.Add(8, 2, 10, 3, 2);
            rd.InsertRow(1, 2);
            Assert.AreEqual(0, rd[2, 2]);
            Assert.AreEqual(1, rd[3, 3]);
            Assert.AreEqual(1, rd[7, 1]);
            Assert.AreEqual(0, rd[10, 1]);
            Assert.AreEqual(2, rd[10, 2]);
            Assert.AreEqual(0, rd[12, 1]);
            Assert.AreEqual(2, rd[12, 2]);
        }
        [TestMethod]
        public void VerifyInsertBeforeRow()
        {
            var rd = new RangeDictionary<int>();

            rd.Add(2, 1, 5, 5, 1);
            rd.Add(8, 2, 10, 3, 2);
            rd.InsertRow(1, 2);
            Assert.AreEqual(0, rd[3, 3]);
            Assert.AreEqual(1, rd[4, 4]);
            Assert.AreEqual(1, rd[7, 1]);
            Assert.AreEqual(0, rd[10, 1]);
            Assert.AreEqual(2, rd[10, 2]);
            Assert.AreEqual(0, rd[12, 1]);
            Assert.AreEqual(2, rd[12, 2]);
        }
        [TestMethod]
        public void VerifyInsertRowSingleColumn()
        {
            var rd = new RangeDictionary<int>();

            rd.Add(2, 1, 5, 5, 1);
            rd.Add(8, 2, 10, 3, 2);
            rd.InsertRow(1, 2, 3, 3);
            Assert.AreEqual(1, rd[2, 2]);
            Assert.AreEqual(0, rd[3, 3]);
            Assert.AreEqual(1, rd[4, 4]);
            Assert.AreEqual(0, rd[7, 1]);
            Assert.AreEqual(2, rd[10, 2]);
            Assert.AreEqual(2, rd[11, 3]);
            Assert.AreEqual(0, rd[12, 1]);
            Assert.AreEqual(2, rd[12, 3]);
        }

        [TestMethod]
        public void VerifyDeleteOnRowOneRow()
        {
            var rd = new RangeDictionary<int>();

            rd.Add(1, 1, 5, 5, 1);
            rd.Add(8, 2, 10, 3, 2);

            rd.DeleteRow(1, 2);
            Assert.AreEqual(1, rd[1, 1]);
            Assert.AreEqual(1, rd[2, 2]);
            Assert.AreEqual(1, rd[3, 3]);
            Assert.AreEqual(0, rd[4, 4]);

            //Assert.AreEqual(2, rd[6, 2]);
        }
        [TestMethod]
        public void VerifyDeleteBeforeRowWithDeleteOneRow()
        {
            var rd = new RangeDictionary<int>();

            rd.Add(1, 1, 1, 5, 1);
            rd.Add(4, 1, 6, 5, 2);

            rd.DeleteRow(2, 3);
            Assert.AreEqual(1, rd[1, 1]);
            Assert.AreEqual(2, rd[2, 2]);
            Assert.AreEqual(2, rd[3, 3]);
            Assert.AreEqual(0, rd[4, 4]);

            //Assert.AreEqual(2, rd[6, 2]);
        }
    }
}
