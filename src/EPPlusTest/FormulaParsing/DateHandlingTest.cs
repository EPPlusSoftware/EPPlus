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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class DateHandlingTest
    {
        [TestMethod]
        public void DateFunctionsWorkWithDifferentCultureDateFormats_US()
        {
            var currentCulture = CultureInfo.CurrentCulture;
#if Core
            var us = CultureInfo.DefaultThreadCurrentCulture = new CultureInfo("en-US");
#else
            var us = CultureInfo.CreateSpecificCulture("en-US");
            Thread.CurrentThread.CurrentCulture = us;
#endif
            double usEoMonth = 0d, usEdate = 0d;
            var thread = new Thread(delegate ()
            {
                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Sheet1");
                    ws.Cells[2, 2].Value = "1/15/2014";
                    ws.Cells[3, 3].Formula = "EOMONTH(C2, 0)";
                    ws.Cells[2, 3].Formula = "EDATE(B2, 0)";
                    ws.Calculate();
                    usEoMonth = Convert.ToDouble(ws.Cells[2, 3].Value);
                    usEdate = Convert.ToDouble(ws.Cells[3, 3].Value);

                }
            });
            thread.Start();
            thread.Join();
            Assert.AreEqual(41654.0, usEoMonth);
            Assert.AreEqual(41670.0, usEdate);
#if Core
            CultureInfo.DefaultThreadCurrentCulture = currentCulture;
#else
            Thread.CurrentThread.CurrentCulture = currentCulture;
#endif
        }

        [TestMethod]
        public void DateFunctionsWorkWithDifferentCultureDateFormats_GB()
        {
            var currentCulture = CultureInfo.CurrentCulture;

            double gbEoMonth = 0d, gbEdate = 0d;
            var thread = new Thread(delegate ()
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-GB");
                using (var package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("Sheet1");
                    ws.Cells[2, 2].Value = "15/1/2014";
                    ws.Cells[3, 3].Formula = "EOMONTH(C2, 0)";
                    ws.Cells[2, 3].Formula = "EDATE(B2, 0)";
                    ws.Calculate();
                    try
                    {
                        gbEoMonth = Convert.ToDouble(ws.Cells[2, 3].Value);
                        gbEdate = Convert.ToDouble(ws.Cells[3, 3].Value);
                    }
                    catch (Exception ex)
                    {
                        Assert.Fail($"Fail culture {Thread.CurrentThread.CurrentCulture.Name}\r\n{ex.Message}\r\n{ex.StackTrace}");
                    }
                }
            });
            thread.Start();
            thread.Join();
            Assert.AreEqual(41654.0, gbEoMonth);
            Assert.AreEqual(41670.0, gbEdate);
#if Core
            CultureInfo.DefaultThreadCurrentCulture = currentCulture;
#else
            Thread.CurrentThread.CurrentCulture = currentCulture;
#endif
        }

        #region Date1904 Test Cases
        [TestMethod]
        public void TestDate1904WithoutSetting()
        {
            var dt1 = new DateTime(2008, 2, 29);
            var dt2 = new DateTime(1950, 11, 30);

            ExcelPackage pck = new ExcelPackage();
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("test");
            ws.Cells[1, 1, 2, 1].Style.Numberformat.Format = ExcelNumberFormat.GetFromBuildInFromID(14);
            ws.Cells[1, 1].Value = dt1;
            ws.Cells[2, 1].Value = dt2;
            pck.Save();


            var pck2 = new ExcelPackage(pck.Stream);
            var ws2 = pck2.Workbook.Worksheets["test"];

            Assert.AreEqual(dt1, ws2.Cells[1, 1].Value);
            Assert.AreEqual(dt2, ws2.Cells[2, 1].Value);

            pck.Dispose();
            pck2.Dispose();
        }

        [TestMethod]
        public void TestDate1904WithSetting()
        {
            var dt1 = new DateTime(2008, 2, 29);
            var dt2 = new DateTime(1950, 11, 30);

            ExcelPackage pck = new ExcelPackage();
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("test");
            pck.Workbook.Date1904 = false;
            ws.Cells[1, 1, 2, 1].Style.Numberformat.Format = ExcelNumberFormat.GetFromBuildInFromID(14);
            ws.Cells[1, 1].Value = dt1;
            ws.Cells[2, 1].Value = dt2;
            pck.Save();

            var pck2 = new ExcelPackage(pck.Stream);
            var ws2 = pck2.Workbook.Worksheets["test"];

            Assert.AreEqual(dt1, ws2.Cells[1, 1].Value);
            Assert.AreEqual(dt2, ws2.Cells[2, 1].Value);

            pck.Dispose();
            pck2.Dispose();
        }

        [TestMethod]
        public void TestDate1904SetAndRemoveSetting()
        {
            var dt1 = new DateTime(2008, 2, 29);
            var dt2 = new DateTime(1950, 11, 30);

            ExcelPackage pck = new ExcelPackage();
            pck.Workbook.Date1904 = true;
            var ws = pck.Workbook.Worksheets.Add("test");
            ws.Cells[1, 1, 2, 1].Style.Numberformat.Format = ExcelNumberFormat.GetFromBuildInFromID(14);
            ws.Cells[1, 1].Value = dt1;
            ws.Cells[2, 1].Value = dt2;
            pck.Save();


            var pck2 = new ExcelPackage(pck.Stream);
            pck2.Workbook.Date1904 = false;
            pck2.Save();

            var pck3 = new ExcelPackage(pck2.Stream);
            ExcelWorksheet ws3 = pck3.Workbook.Worksheets["test"];

            Assert.AreEqual(dt1.AddDays(365.5 * -4), ws3.Cells[1, 1].Value);
            Assert.AreEqual(dt2.AddDays(365.5 * -4), ws3.Cells[2, 1].Value);

            pck.Dispose();
            pck2.Dispose();
            pck3.Dispose();
        }

        [TestMethod]
        public void TestDate1904SetAndSetSetting()
        {
            var dt1 = new DateTime(2008, 2, 29);
            var dt2 = new DateTime(1950, 11, 30);

            ExcelPackage pck = new ExcelPackage();
            pck.Workbook.Date1904 = true;

            var ws = pck.Workbook.Worksheets.Add("test");
            ws.Cells[1, 1, 2, 1].Style.Numberformat.Format = ExcelNumberFormat.GetFromBuildInFromID(14);
            ws.Cells[1, 1].Value = dt1;
            ws.Cells[2, 1].Value = dt2;
            pck.Save();


            var pck2 = new ExcelPackage(pck.Stream);
            pck2.Workbook.Date1904 = true;  // Only the cells must be updated when this change, if set the same nothing must change
            pck2.Save();


            var pck3 = new ExcelPackage(pck2.Stream);
            var ws3 = pck3.Workbook.Worksheets["test"];

            Assert.AreEqual(dt1, ws3.Cells[1, 1].Value);
            Assert.AreEqual(dt2, ws3.Cells[2, 1].Value);
            pck.Dispose();
            pck2.Dispose();
            pck3.Dispose();
        }
        #endregion
    }
}
