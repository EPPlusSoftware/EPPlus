using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ChartExTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            //_pck = OpenPackage("ErrorBars.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            //SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void ReadChartEx()
        {
            using (var p=OpenTemplatePackage("Chartex.xlsx"))
            {
                var drawing = p.Workbook.Worksheets[0].Drawings[0];
                Assert.IsNotNull(((ExcelChartEx)drawing).Fill);
            }
        }
    }
}
