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
using System.Net.Http.Headers;

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
            using(var p=new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                LoadTestdata(ws, 100, 1,1,false,false,new DateTime(2022,11,1));

                p.Workbook.Names.AddFormula("SumOfSheet1", "Sum(Sheet1!$A$2:$A$10)");
                ws.Cells["L1"].Formula = "Sheet1!$B$2+SumOfSheet1+15";
                ws.Calculate();
                var ie = ws.IgnoredErrors.Add(ws.Cells["A1"]);                
                Assert.AreEqual(403830D, p.Workbook.Names["SumOfSheet1"].Value);
                Assert.AreEqual(403847D, ws.Cells["L1"].Value);
            }
        }
        [TestMethod]
        public void VerifyDefinedNameWithMultipleAddresses_Sum()
        {
            using (var p = OpenPackage("NameMultipleAddresses.xlsx"))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 1;
                ws.Cells["A2"].Value = 2;
                ws.Names.Add("MultAddr", ws.Cells["A1,A2"]);
                ws.Cells["A3"].Formula = "Sum(MultAddr)";
                ws.Calculate();

                Assert.AreEqual(3D, ws.Cells["A3"].Value);
            }
        }
        [TestMethod]
        public void VerifyWithMultipleAddresses_Sum()
        {
            using (var p = OpenPackage("NameMultipleAddresses.xlsx"))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 1;
                ws.Cells["A2"].Value = 2;
                ws.Cells["A3"].Value = 3;
                ws.Cells["A4"].Formula = "Sum((A1,A2),A3)";
                ws.Calculate();

                Assert.AreEqual(6D, ws.Cells["A3"].Value);
            }
        }

        [TestMethod]
        public void VerifyDefinedNameWithMultipleAddresses_Avg()
        {
            using (var p = OpenPackage("NameMultipleAddresses.xlsx"))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                ws.Cells["A1"].Value = 1;
                ws.Cells["A2"].Value = 2;
                ws.Cells["A3"].Value = 3;
                ws.Cells["A4"].Value = 4;
                ws.Cells["A5"].Value = 5;
                ws.Names.Add("MultAddr", ws.Cells["A1,A2,A4:A5"]);
                ws.Cells["A6"].Formula = "Average(MultAddr)";
                ws.Calculate();

                Assert.AreEqual(3D, ws.Cells["A6"].Value);
            }
        }
        [TestMethod]
        public void ReadRelativeAddressesInDefinedName()
        {
            using(var p = OpenTemplatePackage("DefinedNameRelative.xlsx"))
            {
                var ws0 = p.Workbook.Worksheets[0];
                var ws1 = p.Workbook.Worksheets[1];
                ws0.ClearFormulaValues();
                ws1.ClearFormulaValues();
                
                p.Workbook.Calculate();

                //Check dynamic array
                Assert.AreEqual(0D, ws0.Cells["F6"].Value);
                Assert.AreEqual(0D, ws0.Cells["F10"].Value);
                Assert.IsNull(ws0.Cells["F11"].Value);

                Assert.AreEqual(3D, ws0.Cells["I9"].Value);
                Assert.AreEqual(3D, ws0.Cells["I10"].Value);
                Assert.AreEqual(3D, ws0.Cells["I11"].Value);
                
                Assert.AreEqual(5D, ws0.Cells["K11"].Value);
                Assert.AreEqual("L11", ws0.Cells["M11"].Value);

                Assert.AreEqual(1D, ws0.Cells["I12"].Value); //RelativeRow
                Assert.AreEqual(3D, ws0.Cells["I16"].Value); //RelativeRow

                //Worksheet 2 - Names containing Table references.
                Assert.AreEqual(3D, ws1.Cells["D2"].Value); //Table referece #this row
                Assert.AreEqual(9D, ws1.Cells["D3"].Value); //Table referece #this row
                Assert.AreEqual(15D, ws1.Cells["D4"].Value); //Table referece #this row

                Assert.AreEqual(3D, ws1.Cells["L2"].Value); //Table referece #this row
                Assert.AreEqual(9D, ws1.Cells["L3"].Value); //Table referece #this row
                Assert.AreEqual(15D, ws1.Cells["L4"].Value); //Table referece #this row


            }
        }
    }
}
