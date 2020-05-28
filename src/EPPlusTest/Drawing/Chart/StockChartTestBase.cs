using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Drawing.Chart
{
    public abstract class StockChartTestBase : TestBase
    {
        protected void LoadStockChartDataPeriod(ExcelWorksheet ws, MemberInfo[] members=null)
        {
            var l = new List<PeriodData>()
            {
                new PeriodData{ Date=new DateTime(2019,12,31), OpeningPrice=100, HighPrice=100, LowPrice=99, ClosePrice=99.5, Volume=10 },
                new PeriodData{ Date=new DateTime(2020,01,01), OpeningPrice=99.5,HighPrice=102, LowPrice=99, ClosePrice=101, Volume=7 },
                new PeriodData{ Date=new DateTime(2020,01,02), OpeningPrice=101,HighPrice=101, LowPrice=92, ClosePrice=94, Volume=8 },
                new PeriodData{ Date=new DateTime(2020,01,03), OpeningPrice=94,HighPrice=97, LowPrice=93, ClosePrice=96.5, Volume=5 },
                new PeriodData{ Date=new DateTime(2020,01,04), OpeningPrice=99.6,HighPrice=107, LowPrice=96.5, ClosePrice=106, Volume=11 },
                new PeriodData{ Date=new DateTime(2020,01,05), OpeningPrice=106,HighPrice=106, LowPrice=103, ClosePrice=104, Volume=7 },
            };
            ws.Cells["A1"].LoadFromCollection(l, true, OfficeOpenXml.Table.TableStyles.Medium11, BindingFlags.Public | BindingFlags.Instance, members);
            ws.Cells["A1:A10"].Style.Numberformat.Format = "yyyy-MM-dd";
            ws.Cells["B1:D10"].Style.Numberformat.Format = "#,##0";
        }
        protected void LoadStockChartDataText(ExcelWorksheet ws)
        {
            var l = new List<EquityData>()
            {
                new EquityData{ EquityName="EPPlus Software AB",  OpeningPrice=100, HighPrice=100, LowPrice=99, ClosePrice=99.5, Volume=10 },
                new EquityData{ EquityName="Company A", OpeningPrice=99.5,HighPrice=102, LowPrice=99, ClosePrice=101, Volume=7 },
                new EquityData{ EquityName="Company B", OpeningPrice=101,HighPrice=101, LowPrice=92, ClosePrice=94, Volume=8},
                new EquityData{ EquityName="Company C", OpeningPrice=94,HighPrice=97, LowPrice=93, ClosePrice=96.5, Volume=5 },
                new EquityData{ EquityName="Company D", OpeningPrice=99.6,HighPrice=107, LowPrice=96.5, ClosePrice=106, Volume=11 },
                new EquityData{ EquityName="Company F", OpeningPrice=106,HighPrice=106, LowPrice=103, ClosePrice=104, Volume=7 }
            };
            ws.Cells["A1"].LoadFromCollection(l, true, OfficeOpenXml.Table.TableStyles.Medium11);
            ws.Cells["A1:A10"].Style.Numberformat.Format = "yyyy-MM-dd";
            ws.Cells["B1:D10"].Style.Numberformat.Format = "#,##0";
        }
        protected class PeriodData
        {
            public DateTime Date { get; set; }
            public double Volume { get; set; }
            public double OpeningPrice { get; set; }
            public double HighPrice { get; set; }
            public double LowPrice { get; set; }
            public double ClosePrice { get; set; }
        }

        protected class EquityData
        {
            public string EquityName { get; set; }
            public double Volume { get; set; }
            public double OpeningPrice { get; set; }
            public double HighPrice { get; set; }
            public double LowPrice { get; set; }
            public double ClosePrice { get; set; }
        }

    }
}
