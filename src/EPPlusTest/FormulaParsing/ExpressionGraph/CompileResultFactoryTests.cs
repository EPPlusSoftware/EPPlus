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
using System.Globalization;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class CompileResultFactoryTests
    {
#if (!Core)
        [TestMethod]
        public void CalculateUsingEuropeanDates()
        {
            var ci=Thread.CurrentThread.CurrentCulture;
            var us = new CultureInfo("en-US");
            Thread.CurrentThread.CurrentCulture = us;
            var result = CompileResultFactory.Create("1/15/2014");
            var numeric = result.ResultNumeric;
            Assert.AreEqual(41654, numeric);
            var gb = new CultureInfo("en-GB");
            Thread.CurrentThread.CurrentCulture = gb;
            var euroResult = CompileResultFactory.Create("15/1/2014");
            var eNumeric = euroResult.ResultNumeric;
            Assert.AreEqual(41654, eNumeric);
            Thread.CurrentThread.CurrentCulture = ci;
        }
#endif
    }
}
