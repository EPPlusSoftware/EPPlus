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
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ExcelStockChartTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("StockHLC.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void AddStockHLC()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockHLC");
            AddHLCData(ws);
            var chart = ws.Drawings.AddStockChart("StockHLC", eStockChartType.StockHLC, ws.Cells["A1:A7"], ws.Cells["B1:B7"], ws.Cells["C1:C7"], ws.Cells["D1:D7"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
        }

        private void AddHLCData(ExcelWorksheet ws)
        {
            var l = new List<EquityData>()
            {
                new EquityData{ Date=new DateTime(2019,12,31), HighPrice=100, LowPrice=99, ClosePrice=99.5 },
                new EquityData{ Date=new DateTime(2020,01,01), HighPrice=102, LowPrice=99, ClosePrice=101 },
                new EquityData{ Date=new DateTime(2020,01,02), HighPrice=101, LowPrice=92, ClosePrice=94 },
                new EquityData{ Date=new DateTime(2020,01,03), HighPrice=97, LowPrice=93, ClosePrice=96.5},
                new EquityData{ Date=new DateTime(2020,01,04), HighPrice=107, LowPrice=96.5, ClosePrice=106 },
                new EquityData{ Date=new DateTime(2020,01,05), HighPrice=106, LowPrice=103, ClosePrice=104 }
            };
            ws.Cells["A1"].LoadFromCollection(l, true, OfficeOpenXml.Table.TableStyles.Medium11);
        }
        private class EquityData
        {
            public DateTime Date { get; set; }
            public double HighPrice { get; set; }
            public double LowPrice { get; set; }
            public double ClosePrice { get; set; }
        }

    }
}
