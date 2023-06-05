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
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class NameValueProviderTests : TestBase
    {
        //private ExcelDataProvider _excelDataProvider;

        //[TestInitialize]
        //public void Setup()
        //{
        //    _excelDataProvider = MockRepository.GenerateMock<ExcelDataProvider>();
        //}

        //[TestMethod]
        //public void IsNamedValueShouldReturnTrueIfKeyIsANamedValue()
        //{
        //    var dict = new Dictionary<string, object>();
        //    dict.Add("A", "B");
        //    _excelDataProvider.Stub(x => x.GetWorkbookNameValues())
        //        .Return(dict);
        //    var nameValueProvider = new EpplusNameValueProvider(_excelDataProvider);

        //    var result = nameValueProvider.IsNamedValue("A");
        //    Assert.IsTrue(result);
        //}

        //[TestMethod]
        //public void IsNamedValueShouldReturnFalseIfKeyIsNotANamedValue()
        //{
        //    var dict = new Dictionary<string, object>();
        //    dict.Add("A", "B");
        //    _excelDataProvider.Stub(x => x.GetWorkbookNameValues())
        //        .Return(dict);
        //    var nameValueProvider = new EpplusNameValueProvider(_excelDataProvider);

        //    var result = nameValueProvider.IsNamedValue("C");
        //    Assert.IsFalse(result);
        //}

        //[TestMethod]
        //public void GetNamedValueShouldReturnCorrectValueIfKeyExists()
        //{
        //    var dict = new Dictionary<string, object>();
        //    dict.Add("A", "B");
        //    _excelDataProvider.Stub(x => x.GetWorkbookNameValues())
        //        .Return(dict);
        //    var nameValueProvider = new EpplusNameValueProvider(_excelDataProvider);

        //    var result = nameValueProvider.GetNamedValue("A");
        //    Assert.AreEqual("B", result);
        //}

        //[TestMethod]
        //public void ReloadShouldReloadDataFromExcelDataProvider()
        //{
        //    var dict = new Dictionary<string, object>();
        //    dict.Add("A", "B");
        //    _excelDataProvider.Stub(x => x.GetWorkbookNameValues())
        //        .Return(dict);
        //    var nameValueProvider = new EpplusNameValueProvider(_excelDataProvider);

        //    var result = nameValueProvider.GetNamedValue("A");
        //    Assert.AreEqual("B", result);

        //    dict.Clear();
        //    nameValueProvider.Reload();
        //    Assert.IsFalse(nameValueProvider.IsNamedValue("A"));
        //}

        [TestMethod]
        public void CalculateWorkbookNameFormula()
        {
            using(var p=OpenPackage("NameWorkbook"))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                LoadTestdata(ws);

                p.Workbook.Names.AddFormula("SumOfSheet1", "Sum(Sheet1!A2:A10)");
                ws.Cells["L1"].Formula = "Sheet1!B2+SumOfSheet1+15";
                ws.Calculate();

                Assert.AreEqual(403830D, p.Workbook.Names["SumOfSheet1"].Value);
                Assert.AreEqual(403847D, ws.Cells["L1"].Value);
            }
        }

    }
}
