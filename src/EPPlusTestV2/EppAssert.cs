using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest
{
    internal class EppAssert
    {
        public static void DoublesAreEqual(double expected, object actual, int nRoundingDecimals)
        {
            if (actual == null) throw new AssertFailedException("EppAssert.DoublesAreEqual failed: parameter actual was null");
            if (!(actual is double d)) throw new AssertFailedException("EppAssert.DoublesAreEqual failed: parameter actual is not a double");
            var roundedResult = Math.Round(d, nRoundingDecimals);
            Assert.AreEqual(expected, roundedResult);
        }

        public static void DoublesAreEqual(double expected, object actual, int nRoundingDecimals, string msg)
        {
            if (actual == null) throw new AssertFailedException("EppAssert.DoublesAreEqual failed: parameter actual was null. " + msg);
            if (!(actual is double d)) throw new AssertFailedException("EppAssert.DoublesAreEqual failed: parameter actual is not a double. " + msg);
            var roundedResult = Math.Round(d, nRoundingDecimals);
            Assert.AreEqual(expected, roundedResult, msg);
        }
    }
}
