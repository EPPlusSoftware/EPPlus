/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class ArgumentParsersImplementationsTests
    {
        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void IntParserShouldThrowIfArgumentIsNull()
        {
            var parser = new IntArgumentParser();
            parser.Parse(null);
        }

        [TestMethod]
        public void IntParserShouldConvertToAnInteger()
        {
            var parser = new IntArgumentParser();
            var result = parser.Parse(3);
            Assert.AreEqual(3, result);
        }

        [TestMethod]
        public void IntParserShouldConvertADoubleToAnInteger()
        {
            var parser = new IntArgumentParser();
            var result = parser.Parse(3d);
            Assert.AreEqual(3, result);
        }

        [TestMethod]
        public void IntParserShouldConvertAStringValueToAnInteger()
        {
            var parser = new IntArgumentParser();
            var result = parser.Parse("3");
            Assert.AreEqual(3, result);
        }

        [TestMethod]
        public void BoolParserShouldConvertNullToFalse()
        {
            var parser = new BoolArgumentParser();
            var result = (bool)parser.Parse(null);
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void BoolParserShouldConvertStringValueTrueToBoolValueTrue()
        {
            var parser = new BoolArgumentParser();
            var result = (bool)parser.Parse("true");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void BoolParserShouldConvert0ToFalse()
        {
            var parser = new BoolArgumentParser();
            var result = (bool)parser.Parse(0);
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void BoolParserShouldConvert1ToTrue()
        {
            var parser = new BoolArgumentParser();
            var result = (bool)parser.Parse(0);
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void DoubleParserShouldConvertDoubleToDouble()
        {
            var parser = new DoubleArgumentParser();
            var result = parser.Parse(3d);
            Assert.AreEqual(3d, result);
        }

        [TestMethod]
        public void DoubleParserShouldConvertIntToDouble()
        {
            var parser = new DoubleArgumentParser();
            var result = parser.Parse(3);
            Assert.AreEqual(3d, result);
        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void DoubleParserShouldThrowIfArgumentIsNull()
        {
            var parser = new DoubleArgumentParser();
            parser.Parse(null);
        }

        [TestMethod]
        public void DoubleParserConvertStringToDoubleWithDotSeparator()
        {
            var parser = new DoubleArgumentParser();
            var result = parser.Parse("3.3");
            Assert.AreEqual(3.3d, result);
        }

        [TestMethod]
        public void DoubleParserConvertDateStringToDouble()
        {
            var parser = new DoubleArgumentParser();
            var result = parser.Parse("3.3.2015");
            Assert.AreEqual(new DateTime(2015,3,3).ToOADate(), result);
        }
    }
}
