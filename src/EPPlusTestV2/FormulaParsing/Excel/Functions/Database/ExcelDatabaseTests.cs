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
    public class ExcelDatabaseTests
    {
        [TestMethod]
        public void DatabaseShouldReadFields()
        {
            using (var package = new ExcelPackage())
            {
                var database = GetDatabase(package);

                Assert.AreEqual(2, database.Fields.Count(), "count was not 2");
                Assert.AreEqual("col1", database.Fields.First().FieldName, "first fieldname was not 'col1'");
                Assert.AreEqual("col2", database.Fields.Last().FieldName, "last fieldname was not 'col12'");
            }
        }

        [TestMethod]
        public void HasMoreRowsShouldBeTrueWhenInitialized()
        {
            using (var package = new ExcelPackage())
            {
                var database = GetDatabase(package);

                Assert.IsTrue(database.HasMoreRows);
            }
            
        }

        [TestMethod]
        public void HasMoreRowsShouldBeFalseWhenLastRowIsRead()
        {
            using (var package = new ExcelPackage())
            {
                var database = GetDatabase(package);
                database.Read();

                Assert.IsFalse(database.HasMoreRows);
            }

        }

        [TestMethod]
        public void DatabaseShouldReadFieldsInRow()
        {
            using (var package = new ExcelPackage())
            {
                var database = GetDatabase(package);
                var row = database.Read();

                Assert.AreEqual(1, row["col1"]);
                Assert.AreEqual(2, row["col2"]);
            }

        }

        private static ExcelDatabase GetDatabase(ExcelPackage package)
        {
            var provider = new EpplusExcelDataProvider(package);
            var sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Value = "col1";
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["B1"].Value = "col2";
            sheet.Cells["B2"].Value = 2;
            var database = new ExcelDatabase(provider, "A1:B2");
            return database;
        }
    }
}
