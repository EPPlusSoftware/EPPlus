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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing.IntegrationTests.ExcelDataProviderTests
{
    [TestClass]
    public class ExcelDataProviderIntegrationTests
    {
        private ExcelCell CreateItem(object val, int row)
        {
            return new ExcelCell(val, null, 0, row);
        }

      

        //[TestMethod]
        //public void ShouldExecuteFormulaInRange()
        //{
        //    var expectedAddres = "A1:A2";
        //    var provider = MockRepository.GenerateStub<ExcelDataProvider>();
        //    provider
        //        .Stub(x => x.GetRangeValues(expectedAddres))
        //        .Return(new object[] { 1, new ExcelCell(null, "SUM(1,2)", 0, 1) });
        //    var parser = new FormulaParser(provider);
        //    var result = parser.Parse(string.Format("sum({0})", expectedAddres));
        //    Assert.AreEqual(4d, result);
        //}

        //[TestMethod, ExpectedException(typeof(CircularReferenceException))]
        //public void ShouldHandleCircularReference2()
        //{
        //    var expectedAddres = "A1:A2";
        //    var provider = MockRepository.GenerateStub<ExcelDataProvider>();
        //    provider
        //        .Stub(x => x.GetRangeValues(expectedAddres))
        //        .Return(new ExcelCell[] { CreateItem(1, 0), new ExcelCell(null, "SUM(A1:A2)",0, 1) });
        //    var parser = new FormulaParser(provider);
        //    var result = parser.Parse(string.Format("sum({0})", expectedAddres));
        //}
    }
}
