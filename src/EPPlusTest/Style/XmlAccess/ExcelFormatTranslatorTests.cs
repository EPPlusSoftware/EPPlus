using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.Style.XmlAccess
{
    [TestClass]
    public class ExcelFormatTranslatorTests
    {
        [TestMethod]
        public void ShouldHandleInvalidStringWithinBrackets()
        {
            var translator = new ExcelFormatTranslator("[YYYY-MM]", -1);
            var r = translator.GetFormatPart(45677);
            Assert.IsFalse(r.IsValid);
        }


        [TestMethod]
        public void ShouldAcceptValidConditionWithDecimal()
        {
            var translator = new ExcelFormatTranslator("[<=-123.45]", -1);
            var r = translator.GetFormatPart(123);
            Assert.IsTrue(r.IsValid);
        }


        [TestMethod]
        public void ShouldHandleValidStringWithinBrackets_1()
        {
            var translator = new ExcelFormatTranslator("[Red]", -1);
            var r = translator.GetFormatPart(45677);
            Assert.IsTrue(r.IsValid);
            Assert.AreEqual("", r.NetFormat);
        }

        [TestMethod]
        public void ShouldHandleValidStringWithinBrackets_2()
        {
            var translator = new ExcelFormatTranslator("[mm]:ss", -1);
            var r = translator.GetFormatPart(45677);
            Assert.IsTrue(r.IsValid);
            Assert.AreEqual("[m]:ss", r.NetFormat);
        }

        [TestMethod]
        public void ShouldHandleValidStringWithinBrackets_3()
        {
            var translator = new ExcelFormatTranslator("[$USD-409]", -1);
            var r = translator.GetFormatPart(45677);
            Assert.IsTrue(r.IsValid);
            Assert.AreEqual("\"USD\"", r.NetFormat);
        }

        [TestMethod]
        public void ShouldHandleIndexedColor()
        {
            var translator = new ExcelFormatTranslator("[Color12]", -1);
            var r = translator.GetFormatPart(123);
            Assert.IsTrue(r.IsValid);
        }


        [TestMethod]
        public void ShouldRejectConditionWithDoubleDecimal()
        {
            var translator = new ExcelFormatTranslator("[>12.3.4]", -1);
            var r = translator.GetFormatPart(123);
            Assert.IsFalse(r.IsValid);
        }

        [TestMethod]
        public void ShouldRejectConditionWithInvalidCharacters()
        {
            var translator = new ExcelFormatTranslator("[>12a3]", -1);
            var r = translator.GetFormatPart(123);
            Assert.IsFalse(r.IsValid);
        }

        [TestMethod]
        public void ShouldRejectConditionWithOnlyOperator()
        {
            var translator = new ExcelFormatTranslator("[>]", -1);
            var r = translator.GetFormatPart(123);
            Assert.IsFalse(r.IsValid);
        }



    }
}