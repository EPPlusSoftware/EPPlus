using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
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
        public void VerifyMergeAddress1()
        {
            var rd = new RangeDictionary();

            var r1 = new FormulaRangeAddress() { FromRow = 1, ToRow = 5, FromCol = 1, ToCol = 5 };
            var b1=rd.Merge(ref r1);
            var r2=new FormulaRangeAddress() { FromRow = 6, ToRow = 10, FromCol = 1, ToCol = 3 };
            var b2 = rd.Merge(ref r2);
            var r3 = new FormulaRangeAddress() { FromRow = 1, ToRow = 2, FromCol = 1, ToCol = 3 };
            var b3 = rd.Merge(ref r3);
            var r4 = new FormulaRangeAddress() { FromRow = 7, ToRow = 8, FromCol = 1, ToCol = 3 };
            var b4 = rd.Merge(ref r4);
            var r5 = new FormulaRangeAddress() { FromRow = 3, ToRow = 8, FromCol = 3, ToCol = 6 };
            var b5 = rd.Merge(ref r5);

            Assert.AreEqual(9, rd._addresses[1][0]);
            Assert.AreEqual(9, rd._addresses[2][0]);
            Assert.AreEqual(9, rd._addresses[3][0]);
            Assert.AreEqual(7, rd._addresses[4][0]);
            Assert.AreEqual(7, rd._addresses[5][0]);
            Assert.AreEqual(2097159, rd._addresses[6][0]);

            Assert.IsTrue(rd.Exists(1, 1));
            Assert.IsTrue(rd.Exists(6, 3));
            Assert.IsTrue(rd.Exists(7, 3));
            Assert.IsFalse(rd.Exists(2, 6));
            Assert.IsTrue(rd.Exists(3, 6));
            Assert.IsTrue(rd.Exists(8, 6));
            Assert.IsFalse(rd.Exists(9, 6));
        }
        [TestMethod]
        public void VerifyMergeAddressBetween()
        {
            var rd = new RangeDictionary();

            var r1 = new FormulaRangeAddress() { FromRow = 2, ToRow = 5, FromCol = 1, ToCol = 5 };
            var b1 = rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 7, ToRow = 10, FromCol = 1, ToCol = 3 };
            var b2 = rd.Merge(ref r2);
            var r3 = new FormulaRangeAddress() { FromRow = 6, ToRow = 6, FromCol = 1, ToCol = 3 };
            var b3 = rd.Merge(ref r3);
            var r4 = new FormulaRangeAddress() { FromRow = 1, ToRow = 1, FromCol = 1, ToCol = 10 };
            var b4 = rd.Merge(ref r4);

            Assert.IsTrue(rd.Exists(1, 1));
        }
        [TestMethod]
        public void VerifyOrderIsSorted()
        {
            var rd = new RangeDictionary();

            var r1 = new FormulaRangeAddress() { FromRow = 7, ToRow = 7, FromCol = 1, ToCol = 5 };
            rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 5, ToRow = 5, FromCol = 1, ToCol = 5 };
            rd.Merge(ref r2);
            var r3 = new FormulaRangeAddress() { FromRow = 15, ToRow = 15, FromCol = 1, ToCol = 5 };
            rd.Merge(ref r3);
            var r4 = new FormulaRangeAddress() { FromRow = 9, ToRow = 9, FromCol = 1, ToCol = 5 };
            rd.Merge(ref r4);
            var r5 = new FormulaRangeAddress() { FromRow = 11, ToRow = 11, FromCol = 1, ToCol = 5 };
            rd.Merge(ref r5);
            var r6 = new FormulaRangeAddress() { FromRow = 1, ToRow = 2, FromCol = 1, ToCol = 5 };
            rd.Merge(ref r6);
            var r7 = new FormulaRangeAddress() { FromRow = 13, ToRow = 13, FromCol = 1, ToCol = 5 };
            rd.Merge(ref r7);

            Assert.AreEqual(1, GetFromRow(rd._addresses[1][0]));
            Assert.AreEqual(5, GetFromRow(rd._addresses[1][1]));
            Assert.AreEqual(7, GetFromRow(rd._addresses[1][2]));
            Assert.AreEqual(9, GetFromRow(rd._addresses[1][3]));
            Assert.AreEqual(11, GetFromRow(rd._addresses[1][4]));
            Assert.AreEqual(13, GetFromRow(rd._addresses[1][5]));
            Assert.AreEqual(15, GetFromRow(rd._addresses[1][6]));

            Assert.IsTrue(rd.Exists(1, 1));
        }


        [TestMethod]
        public void VerifySpillRangesMergeFrom()
        {
            var rd = new RangeDictionary();
            var r1 = new FormulaRangeAddress() { FromRow = 7, ToRow = 18, FromCol = 1, ToCol = 5 };
            rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 5, ToRow = 18, FromCol = 1, ToCol = 5 };
            var b2=rd.Merge(ref r2);

            Assert.AreEqual(5, GetFromRow(rd._addresses[1][0]));
            Assert.AreEqual(18, GetToRow(rd._addresses[1][0]));
        }
        [TestMethod]
        public void VerifySpillRangesMergeTo()
        {
            var rd = new RangeDictionary();
            var r1 = new FormulaRangeAddress() { FromRow = 5, ToRow = 18, FromCol = 1, ToCol = 5 };
            var b1 = rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 5, ToRow = 21, FromCol = 1, ToCol = 5 };
            var b2 = rd.Merge(ref r2);

            Assert.AreEqual(5, GetFromRow(rd._addresses[1][0]));
            Assert.AreEqual(21, GetToRow(rd._addresses[1][0]));
            Assert.IsTrue(b1);
            Assert.IsTrue(b2);
        }
        [TestMethod]
        public void VerifyRangesMergeInto()
        {
            var rd = new RangeDictionary();
            var r1 = new FormulaRangeAddress() { FromRow = 7, ToRow = 18, FromCol = 1, ToCol = 5 };
            var b1=rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 5, ToRow = 21, FromCol = 1, ToCol = 5 };
            var b2 = rd.Merge(ref r2);

            Assert.AreEqual(5, GetFromRow(rd._addresses[1][0]));
            Assert.AreEqual(21, GetToRow(rd._addresses[1][0]));
            Assert.IsTrue(b1);
            Assert.IsTrue(b2);
        }
        [TestMethod]
        public void VerifyMergeRangesSame()
        {
            var rd = new RangeDictionary();
            var r1 = new FormulaRangeAddress() { FromRow = 7, ToRow = 18, FromCol = 1, ToCol = 5 };
            var b1 = rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 7, ToRow = 18, FromCol = 1, ToCol = 5 };
            var b2 = rd.Merge(ref r2);

            Assert.AreEqual(7, GetFromRow(rd._addresses[1][0]));
            Assert.AreEqual(18, GetToRow(rd._addresses[1][0]));
            Assert.IsTrue(b1);
            Assert.IsFalse(b2);
        }
        [TestMethod]
        public void VerifyMergeRangesTopLeft()
        {
            var rd = new RangeDictionary();
            var r1 = new FormulaRangeAddress() { FromRow = 5, ToRow = 12, FromCol = 2, ToCol = 5 };
            var b1 = rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 1, ToRow = 5, FromCol = 1, ToCol = 5 };
            var b2 = rd.Merge(ref r2);

            Assert.AreEqual(1, GetFromRow(rd._addresses[2][0]));
            Assert.AreEqual(12, GetToRow(rd._addresses[2][0]));
            Assert.IsTrue(b1);
            Assert.IsTrue(b2);
        }
        [TestMethod]
        public void VerifyMergeRangesTop()
        {
            var rd = new RangeDictionary();
            var r1 = new FormulaRangeAddress() { FromRow = 5, ToRow = 12, FromCol = 1, ToCol = 5 };
            var b1 = rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 1, ToRow = 5, FromCol = 1, ToCol = 5 };
            var b2 = rd.Merge(ref r2);

            Assert.AreEqual(1, GetFromRow(rd._addresses[2][0]));
            Assert.AreEqual(12, GetToRow(rd._addresses[2][0]));
            Assert.IsTrue(b1);
            Assert.IsTrue(b2);

            //Rest A1:E4
            Assert.AreEqual(1, r2.FromRow);
            Assert.AreEqual(1, r2.FromCol);
            Assert.AreEqual(4, r2.ToRow);
            Assert.AreEqual(5, r2.ToCol);
        }
        [TestMethod]
        public void VerifyMergeRangesTopRight()
        {
            var rd = new RangeDictionary();
            var r1 = new FormulaRangeAddress() { FromRow = 5, ToRow = 12, FromCol = 2, ToCol = 5 };
            var b1 = rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 1, ToRow = 6, FromCol = 3, ToCol = 6 };
            var b2 = rd.Merge(ref r2);

            Assert.AreEqual(1, GetFromRow(rd._addresses[3][0]));
            Assert.AreEqual(12, GetToRow(rd._addresses[3][0]));
            Assert.IsTrue(b1);
            Assert.IsTrue(b2);

            //Full range A1:E4
            Assert.AreEqual(1, r2.FromRow);
            Assert.AreEqual(3, r2.FromCol);
            Assert.AreEqual(6, r2.ToRow);
            Assert.AreEqual(6, r2.ToCol);
        }
        [TestMethod]
        public void VerifyRangesTopRight()
        {
            var rd = new RangeDictionary();
            var r1 = new FormulaRangeAddress() { FromRow = 5, ToRow = 12, FromCol = 2, ToCol = 5 };
            var b1 = rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 1, ToRow = 6, FromCol = 3, ToCol = 6 };
            var b2 = rd.Merge(ref r2);

            Assert.AreEqual(1, GetFromRow(rd._addresses[3][0]));
            Assert.AreEqual(12, GetToRow(rd._addresses[3][0]));
            Assert.IsTrue(b1);
            Assert.IsTrue(b2);

            //Full range C1:F6
            Assert.AreEqual(1, r2.FromRow);
            Assert.AreEqual(3, r2.FromCol);
            Assert.AreEqual(6, r2.ToRow);
            Assert.AreEqual(6, r2.ToCol);
        }
        [TestMethod]
        public void VerifyRangesLeft()
        {
            var rd = new RangeDictionary();
            var r1 = new FormulaRangeAddress() { FromRow = 5, ToRow = 12, FromCol = 3, ToCol = 5 };
            var b1 = rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 5, ToRow = 12, FromCol = 1, ToCol = 5 };
            var b2 = rd.Merge(ref r2);

            Assert.AreEqual(5, GetFromRow(rd._addresses[3][0]));
            Assert.AreEqual(12, GetToRow(rd._addresses[3][0]));
            Assert.IsTrue(b1);
            Assert.IsTrue(b2);

            //Spill range A5:B12
            Assert.AreEqual(5, r2.FromRow);
            Assert.AreEqual(1, r2.FromCol);
            Assert.AreEqual(12, r2.ToRow);
            Assert.AreEqual(2, r2.ToCol);
        }
        [TestMethod]
        public void VerifyRangesRight()
        {
            var rd = new RangeDictionary();
            var r1 = new FormulaRangeAddress() { FromRow = 5, ToRow = 13, FromCol = 3, ToCol = 5 };
            var b1 = rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 5, ToRow = 13, FromCol = 4, ToCol = 12 };
            var b2 = rd.Merge(ref r2);

            Assert.AreEqual(5, GetFromRow(rd._addresses[4][0]));
            Assert.AreEqual(13, GetToRow(rd._addresses[4][0]));
            Assert.IsTrue(b1);
            Assert.IsTrue(b2);

            //Full range A1:E4
            Assert.AreEqual(5, r2.FromRow);
            Assert.AreEqual(6, r2.FromCol);
            Assert.AreEqual(13, r2.ToRow);
            Assert.AreEqual(12, r2.ToCol);
        }
        [TestMethod]
        public void VerifyRangesBottomLeft()
        {
            var rd = new RangeDictionary();
            var r1 = new FormulaRangeAddress() { FromRow = 5, ToRow = 12, FromCol = 3, ToCol = 5 };
            var b1 = rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 9, ToRow = 14, FromCol = 1, ToCol = 4 };
            var b2 = rd.Merge(ref r2);

            
            Assert.AreEqual(5, GetFromRow(rd._addresses[3][0]));
            Assert.AreEqual(14, GetToRow(rd._addresses[3][0]));
            Assert.IsTrue(b1);
            Assert.IsTrue(b2);

            //Spill range A5:B12
            Assert.AreEqual(9, r2.FromRow);
            Assert.AreEqual(1, r2.FromCol);
            Assert.AreEqual(14, r2.ToRow);
            Assert.AreEqual(4, r2.ToCol);
        }
        [TestMethod]
        public void VerifyRangesBottom()
        {
            var rd = new RangeDictionary();
            var r1 = new FormulaRangeAddress() { FromRow = 5, ToRow = 12, FromCol = 3, ToCol = 5 };
            var b1 = rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 12, ToRow = 14, FromCol = 3, ToCol = 5 };
            var b2 = rd.Merge(ref r2);


            Assert.AreEqual(5, GetFromRow(rd._addresses[3][0]));
            Assert.AreEqual(14, GetToRow(rd._addresses[3][0]));
            Assert.IsTrue(b1);
            Assert.IsTrue(b2);

            //Spill range A5:B12
            Assert.AreEqual(13, r2.FromRow);
            Assert.AreEqual(3, r2.FromCol);
            Assert.AreEqual(14, r2.ToRow);
            Assert.AreEqual(5, r2.ToCol);
        }
        [TestMethod]
        public void VerifyRangesBottomRight()
        {
            var rd = new RangeDictionary();
            var r1 = new FormulaRangeAddress() { FromRow = 5, ToRow = 12, FromCol = 3, ToCol = 5 };
            var b1 = rd.Merge(ref r1);
            var r2 = new FormulaRangeAddress() { FromRow = 9, ToRow = 14, FromCol = 5, ToCol = 6 };
            var b2 = rd.Merge(ref r2);


            Assert.AreEqual(5, GetFromRow(rd._addresses[5][0]));
            Assert.AreEqual(14, GetToRow(rd._addresses[5][0]));
            Assert.IsTrue(b1);
            Assert.IsTrue(b2);

            //Spill range A5:B12
            Assert.AreEqual(9, r2.FromRow);
            Assert.AreEqual(5, r2.FromCol);
            Assert.AreEqual(14, r2.ToRow);
            Assert.AreEqual(6, r2.ToCol);
        }
    }
}
