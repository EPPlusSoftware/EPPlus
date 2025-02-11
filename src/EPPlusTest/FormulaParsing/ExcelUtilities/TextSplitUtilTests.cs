using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExcelUtilities.TextUtils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.ExcelUtilities
{
    [TestClass]
    public class TextSplitUtilTests
    {
        [TestMethod]
        public void GetDelimiterPositions_ShouldFindIndex()
        {
            var util = new TextSplitUtil("abc");
            var positions = util.GetDelimiterPostions("b");
            Assert.AreEqual(1, positions.Length);
            Assert.AreEqual(1, positions[0].Position);
        }

        [TestMethod]
        public void GetDelimiterPositions_ShouldFindIndex_IgnoreCase()
        {
            var util = new TextSplitUtil("abc");
            var positions = util.GetDelimiterPostions("B", true);
            Assert.AreEqual(1, positions.Length);
            Assert.AreEqual(1, positions[0].Position);
        }

        [TestMethod]
        public void GetDelimiterPositions_ShouldFindIndex_MultiCharDelimiter()
        {
            var util = new TextSplitUtil("abc def ghi");
            var positions = util.GetDelimiterPostions("def");
            Assert.AreEqual(1, positions.Length);
            Assert.AreEqual(4, positions[0].Position);
        }

        [TestMethod]
        public void GetDelimiterPositions_ShouldFindIndex_MultiCharDelimiter_MultipleDelimiters()
        {
            var util = new TextSplitUtil("abc def ghi def jkl");
            var positions = util.GetDelimiterPostions("def");
            Assert.AreEqual(2, positions.Length);
            Assert.AreEqual(4, positions[0].Position);
            Assert.AreEqual(12, positions[1].Position);
        }

        [TestMethod]
        public void GetDelimiterPositions_ShouldFindIndex_MultiCharDelimiter_PartlyMatch()
        {
            var util = new TextSplitUtil("abc def ghi def jkl");
            var positions = util.GetDelimiterPostions("dep");
            Assert.AreEqual(0, positions.Length);
        }

        [TestMethod]
        public void GetTextBefore_OneDelimiter()
        {
            var util = new TextSplitUtil("abc");
            var result = util.GetTextBefore("b", 1, false, false, out bool isOutOfRange, out int? matchIndex);
            Assert.IsFalse(isOutOfRange);
            Assert.AreEqual(1, matchIndex);
            Assert.AreEqual("a", result);
        }

        [TestMethod]
        public void GetTextBefore_IgnoreCase()
        {
            var util = new TextSplitUtil("abc");
            var result = util.GetTextBefore("B", 1, true, false, out bool isOutOfRange, out int? matchIndex);
            Assert.IsFalse(isOutOfRange);
            Assert.AreEqual(1, matchIndex);
            Assert.AreEqual("a", result);
        }

        [TestMethod]
        public void GetTextBefore_MultiCharDelimiter()
        {
            var util = new TextSplitUtil("abc def ghi def jkl");
            var result = util.GetTextBefore("def", 1, false, false, out bool isOutOfRange, out int? matchIndex);
            Assert.IsFalse(isOutOfRange);
            Assert.AreEqual(4, matchIndex);
            Assert.AreEqual("abc ", result);
        }

        [TestMethod]
        public void GetTextBefore_MultiCharDelimiter_SecondInstance()
        {
            var util = new TextSplitUtil("abc def ghi def jkl");
            var result = util.GetTextBefore("def", 2, false, false, out bool isOutOfRange, out int? matchIndex);
            Assert.IsFalse(isOutOfRange);
            Assert.AreEqual(12, matchIndex);
            Assert.AreEqual("abc def ghi ", result);
        }

        [TestMethod]
        public void GetTextBefore_MultiCharDelimiter_OutOfRange()
        {
            var util = new TextSplitUtil("abc def ghi def jkl");
            var result = util.GetTextBefore("def", 3, false, false, out bool isOutOfRange, out int? matchIndex);
            Assert.IsTrue(isOutOfRange);
            Assert.IsFalse(matchIndex.HasValue);
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GetTextBefore_MultiCharDelimiter_Reverse()
        {
            var util = new TextSplitUtil("abc def ghi def jkl");
            var result = util.GetTextBefore("def", -1, false, false, out bool isOutOfRange, out int? matchIndex);
            Assert.IsFalse(isOutOfRange);
            Assert.IsTrue(matchIndex.HasValue);
            Assert.AreEqual(12, matchIndex);
            Assert.AreEqual("abc def ghi ", result);
        }

        [TestMethod]
        public void GetTextAfter_OneDelimiter()
        {
            var util = new TextSplitUtil("abc");
            var result = util.GetTextAfter("b", 1, false, false, out bool isOutOfRange, out int? matchIndex);
            Assert.IsFalse(isOutOfRange);
            Assert.AreEqual(1, matchIndex);
            Assert.AreEqual("c", result);
        }

        [TestMethod]
        public void GetTextAfter_IgnoreCase()
        {
            var util = new TextSplitUtil("abc");
            var result = util.GetTextAfter("B", 1, true, false, out bool isOutOfRange, out int? matchIndex);
            Assert.IsFalse(isOutOfRange);
            Assert.AreEqual(1, matchIndex);
            Assert.AreEqual("c", result);
        }

        [TestMethod]
        public void GetTextAfter_MultiCharDelimiter()
        {
            var util = new TextSplitUtil("abc def ghi def jkl");
            var result = util.GetTextAfter("def", 1, false, false, out bool isOutOfRange, out int? matchIndex);
            Assert.IsFalse(isOutOfRange);
            Assert.AreEqual(4, matchIndex);
            Assert.AreEqual(" ghi def jkl", result);
        }

        [TestMethod]
        public void GetTextAfter_MultiCharDelimiter_SecondInstance()
        {
            var util = new TextSplitUtil("abc def ghi def jkl");
            var result = util.GetTextAfter("def", 2, false, false, out bool isOutOfRange, out int? matchIndex);
            Assert.IsFalse(isOutOfRange);
            Assert.AreEqual(12, matchIndex);
            Assert.AreEqual(" jkl", result);
        }

        [TestMethod]
        public void GetTextAfter_MultiCharDelimiter_OutOfRange()
        {
            var util = new TextSplitUtil("abc def ghi def jkl");
            var result = util.GetTextAfter("def", 3, false, false, out bool isOutOfRange, out int? matchIndex);
            Assert.IsTrue(isOutOfRange);
            Assert.IsFalse(matchIndex.HasValue);
            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void GetTextAfter_MultiCharDelimiter_Reverse()
        {
            var util = new TextSplitUtil("abc def ghi def jkl");
            var result = util.GetTextAfter("def", -1, true, false, out bool isOutOfRange, out int? matchIndex);
            Assert.IsFalse(isOutOfRange);
            Assert.AreEqual(12, matchIndex);
            Assert.AreEqual(" jkl", result);
        }

        [TestMethod]
        public void GetTextAfter_MultiCharDelimiter_WithEscapeChar()
        {
            var util = new TextSplitUtil("abc d\"f ghi def jkl");
            var result = util.GetTextAfter("d\"\"f", 1, true, false, out bool isOutOfRange, out int? matchIndex);
            Assert.AreEqual(" ghi def jkl", result);
        }
    }
}
