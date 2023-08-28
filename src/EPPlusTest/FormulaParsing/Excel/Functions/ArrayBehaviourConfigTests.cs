using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestClass]
    public class ArrayBehaviourConfigTests
    {
        [TestMethod]
        public void ShouldReturnCorrectValue_ArrayIntervalOnly()
        {
            var config = new ArrayBehaviourConfig();
            config.ArrayArgInterval = 2;
            var result = config.CanBeArrayArg(0);
            Assert.IsTrue(result);

            result = config.CanBeArrayArg(1);
            Assert.IsFalse(result);

            result = config.CanBeArrayArg(2);
            Assert.IsTrue(result);

            result = config.CanBeArrayArg(3);
            Assert.IsFalse(result);

            config.ArrayArgInterval = 3;
            result = config.CanBeArrayArg(0);
            Assert.IsTrue(result);

            result = config.CanBeArrayArg(1);
            Assert.IsFalse(result);

            result = config.CanBeArrayArg(2);
            Assert.IsFalse(result);

            result = config.CanBeArrayArg(3);
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void ShouldReturnCorrectValue_ArrayIntervalAndIgnoreStart()
        {
            var config = new ArrayBehaviourConfig();
            config.IgnoreNumberOfArgsFromStart = 1;
            config.ArrayArgInterval = 2;
            var result = config.CanBeArrayArg(0);
            Assert.IsFalse(result);

            result = config.CanBeArrayArg(1);
            Assert.IsTrue(result);

            result = config.CanBeArrayArg(2);
            Assert.IsFalse(result);

            result = config.CanBeArrayArg(3);
            Assert.IsTrue(result);

            config.ArrayArgInterval = 3;
            result = config.CanBeArrayArg(0);
            Assert.IsFalse(result);

            result = config.CanBeArrayArg(1);
            Assert.IsTrue(result);

            result = config.CanBeArrayArg(2);
            Assert.IsFalse(result);

            result = config.CanBeArrayArg(3);
            Assert.IsFalse(result);

            result = config.CanBeArrayArg(4);
            Assert.IsTrue(result);
        }
    }
}
