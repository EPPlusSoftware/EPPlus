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
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Database
{
    [TestClass]
    public class CriteriaTests
    {
        [TestMethod]
        public void CriteriaShouldReadFieldsAndValues()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Crit1";
                sheet.Cells["B1"].Value = "Crit2";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = 2;

                var provider = new EpplusExcelDataProvider(package);

                var criteria = new ExcelDatabaseCriteria(provider, "A1:B2");

                Assert.AreEqual(2, criteria.Items.Count);
                Assert.AreEqual("crit1", criteria.Items.Keys.First().ToString());
                Assert.AreEqual("crit2", criteria.Items.Keys.Last().ToString());
                Assert.AreEqual(1, criteria.Items.Values.First());
                Assert.AreEqual(2, criteria.Items.Values.Last());
            }
        }

        [TestMethod]
        public void CriteriaShouldIgnoreEmptyFields1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Crit1";
                sheet.Cells["B1"].Value = "Crit2";
                sheet.Cells["A2"].Value = 1;

                var provider = new EpplusExcelDataProvider(package);

                var criteria = new ExcelDatabaseCriteria(provider, "A1:B2");

                Assert.AreEqual(1, criteria.Items.Count);
                Assert.AreEqual("crit1", criteria.Items.Keys.First().ToString());
                Assert.AreEqual(1, criteria.Items.Values.Last());
            }
        }

        [TestMethod]
        public void CriteriaShouldIgnoreEmptyFields2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Crit1";
                sheet.Cells["A2"].Value = 1;

                var provider = new EpplusExcelDataProvider(package);

                var criteria = new ExcelDatabaseCriteria(provider, "A1:B2");

                Assert.AreEqual(1, criteria.Items.Count);
                Assert.AreEqual("crit1", criteria.Items.Keys.First().ToString());
                Assert.AreEqual(1, criteria.Items.Values.Last());
            }
        }

    }
}
