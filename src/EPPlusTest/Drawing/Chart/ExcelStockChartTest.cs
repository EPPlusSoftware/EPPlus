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
        public void ReadStockVHLC()
        {
            using(var p=OpenTemplatePackage("StockVHLC.xlsx"))
            {
                var c = p.Workbook.Worksheets[0].Drawings[0];
                SaveWorkbook("StockVHLCSaved.xlsx", p);
            }
        }
        [TestMethod]
        public void AddStockHLCText()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockTextHLC");
            AddHLCText(ws);
            
            var chart = ws.Drawings.AddStockChart("StockPeriodHLC", eStockChartType.StockHLC, ws.Cells["A1:A7"], ws.Cells["C1:C7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
        }
        [TestMethod]
        public void AddStockOHLCText()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockTextOHLC");
            AddHLCText(ws);

            var chart = ws.Drawings.AddStockChart("StockTextOHLC", eStockChartType.StockOHLC, ws.Cells["A1:A7"], ws.Cells["C1:C7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"], ws.Cells["B1:B7"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
        }
        [TestMethod]
        public void AddStockHLCPeriod()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockPeriodHLC");
            AddHLCPeriod(ws);

            var chart = ws.Drawings.AddStockChart("StockPeriodHLC", eStockChartType.StockHLC, ws.Cells["A1:A7"], ws.Cells["C1:C7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
        }
        [TestMethod]
        public void AddStockOHLCPeriod()
        {
            var ws = _pck.Workbook.Worksheets.Add("StockPeriodOHLC");
            AddHLCPeriod(ws);

            var chart = ws.Drawings.AddStockChart("StockPeriodOHLC", eStockChartType.StockOHLC, ws.Cells["A1:A7"], ws.Cells["C1:C7"], ws.Cells["D1:D7"], ws.Cells["E1:E7"], ws.Cells["B1:B7"]);
            chart.SetPosition(2, 0, 15, 0);
            chart.SetSize(1600, 900);
        }
        private void AddHLCPeriod(ExcelWorksheet ws)
        {
            var l = new List<PeriodData>()
            {
                new PeriodData{ Date=new DateTime(2019,12,31), OpeningPrice=100, HighPrice=100, LowPrice=99, ClosePrice=99.5 },
                new PeriodData{ Date=new DateTime(2020,01,01), OpeningPrice=99.5,HighPrice=102, LowPrice=99, ClosePrice=101 },
                new PeriodData{ Date=new DateTime(2020,01,02), OpeningPrice=101,HighPrice=101, LowPrice=92, ClosePrice=94 },
                new PeriodData{ Date=new DateTime(2020,01,03), OpeningPrice=94,HighPrice=97, LowPrice=93, ClosePrice=96.5},
                new PeriodData{ Date=new DateTime(2020,01,04), OpeningPrice=99.6,HighPrice=107, LowPrice=96.5, ClosePrice=106 },
                new PeriodData{ Date=new DateTime(2020,01,05), OpeningPrice=106,HighPrice=106, LowPrice=103, ClosePrice=104 }
            };
            ws.Cells["A1"].LoadFromCollection(l, true, OfficeOpenXml.Table.TableStyles.Medium11);
            ws.Cells["A1:A10"].Style.Numberformat.Format = "yyyy-MM-dd";
            ws.Cells["B1:D10"].Style.Numberformat.Format = "#,##0";
        }
        private void AddHLCText(ExcelWorksheet ws)
        {
            var l = new List<EquityData>()
            {
                new EquityData{ EquityName="EPPlus Software AB",OpeningPrice=100, HighPrice=100, LowPrice=99, ClosePrice=99.5 },
                new EquityData{ EquityName="Company A", OpeningPrice=99.5,HighPrice=102, LowPrice=99, ClosePrice=101 },
                new EquityData{ EquityName="Company B", OpeningPrice=101,HighPrice=101, LowPrice=92, ClosePrice=94 },
                new EquityData{ EquityName="Company C", OpeningPrice=94,HighPrice=97, LowPrice=93, ClosePrice=96.5},
                new EquityData{ EquityName="Company D", OpeningPrice=99.6,HighPrice=107, LowPrice=96.5, ClosePrice=106 },
                new EquityData{ EquityName="Company F", OpeningPrice=106,HighPrice=106, LowPrice=103, ClosePrice=104 }
            };
            ws.Cells["A1"].LoadFromCollection(l, true, OfficeOpenXml.Table.TableStyles.Medium11);
            ws.Cells["A1:A10"].Style.Numberformat.Format = "yyyy-MM-dd";
            ws.Cells["B1:D10"].Style.Numberformat.Format = "#,##0";
        }
        private class PeriodData
        {
            public DateTime Date { get; set; }
            public double OpeningPrice { get; set; }
            public double HighPrice { get; set; }
            public double LowPrice { get; set; }
            public double ClosePrice { get; set; }
        }

        private class EquityData
        {
            public string EquityName { get; set; }
            public double OpeningPrice { get; set; }
            public double HighPrice { get; set; }
            public double LowPrice { get; set; }
            public double ClosePrice { get; set; }
        }

    }
}
