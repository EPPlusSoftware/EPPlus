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
  02/11/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;

namespace EPPlusTest.Table.PivotTable
{
    [TestClass]
    public class PivotTableTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
        }
        [ClassCleanup]
        public static void Cleanup()
        {
        }
        [TestMethod]
        public void ValidateLoadSave()
        {
            using (ExcelPackage p1 = new ExcelPackage())
            {
                var tblName = "Table1";
                var tblAddress = "A1:D4";
                var wsData = p1.Workbook.Worksheets.Add("TableData");
                var wsPivot = p1.Workbook.Worksheets.Add("PivotSimple");
                var Table1 = wsData.Tables.Add(wsData.Cells[tblAddress],tblName);
                var pivotTable1 = wsPivot.PivotTables.Add(wsPivot.Cells["A1"], wsData.Cells[Table1.Address.Address], "PivotTable1");

                pivotTable1.RowFields.Add(pivotTable1.Fields[0]);
                pivotTable1.DataFields.Add(pivotTable1.Fields[1]);
                pivotTable1.ColumnFields.Add(pivotTable1.Fields[2]);

                Assert.AreEqual(tblAddress, wsPivot.PivotTables[0].CacheDefinition.SourceRange.Address);
                Assert.AreEqual(Table1.Columns.Count ,pivotTable1.Fields.Count);
                Assert.AreEqual(1, pivotTable1.RowFields.Count);
                Assert.AreEqual(1, pivotTable1.DataFields.Count);
                Assert.AreEqual(1, pivotTable1.ColumnFields.Count);

                p1.Save();

                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    wsData = p2.Workbook.Worksheets[0];
                    wsPivot = p2.Workbook.Worksheets[1];

                    pivotTable1 = wsPivot.PivotTables[0];
                    Assert.AreEqual(tblAddress, pivotTable1.CacheDefinition.SourceRange.Address);
                    Assert.AreEqual(Table1.Columns.Count, pivotTable1.Fields.Count);
                    Assert.AreEqual(1, pivotTable1.RowFields.Count);
                    Assert.AreEqual(1, pivotTable1.DataFields.Count);
                    Assert.AreEqual(1, pivotTable1.ColumnFields.Count);
                }
            }
        }
    }
}
