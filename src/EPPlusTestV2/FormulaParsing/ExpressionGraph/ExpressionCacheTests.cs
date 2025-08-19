using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class ExpressionCacheTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _wsData;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("ExpressionCache.xlsx", true);
            _wsData = _pck.Workbook.Worksheets.Add("Data");
            LoadTestdata(_wsData);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }

        [TestMethod]
        public void CacheExpression()
        {
            var wsFormula = _pck.Workbook.Worksheets.Add("FormulaToCache");
            wsFormula.Cells["A2:A100"].Formula = "Data!D2/Sum(Data!$D$2:$D$100)+1";
            wsFormula.Calculate();
        }
    }
}
