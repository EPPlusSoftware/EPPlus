using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class DependencyChainTests
    {
        [TestMethod]
        public void NoDepth()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                ExcelRangeBase cellA1 = sheet.Cells["A1"];
                cellA1.Formula = "A2";
                IEnumerable<IFormulaCellInfo> depthTree = package.Workbook.FormulaParserManager.GetCalculationChainByDepth(cellA1);

                Assert.AreEqual(depthTree.ToList()[0].Address, cellA1.Address);
            }
        }

        [TestMethod]
        public void LowDepth()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                ExcelRangeBase cellA1 = sheet.Cells["A1"];
                cellA1.Formula = "A2";
                ExcelRangeBase cellA2 = sheet.Cells["A2"];
                cellA2.Formula = "A3";

                IEnumerable<IFormulaCellInfo> depthTree = package.Workbook.FormulaParserManager.GetCalculationChainByDepth(cellA1);

                Assert.AreEqual(depthTree.ToList()[0].Address, cellA1.Address); // A1 must come before A2
                Assert.AreEqual(depthTree.ToList()[1].Address, cellA2.Address); // A2 must be included
            }
        }

        [TestMethod]
        public void BranchAtDepth0()
        {
            using (var package = new ExcelPackage())
            {
                /*
                 *   A1   depth=0
                 *  |  \
                 *  A2  A3  depth=1
                 *  |   |
                 *  B2  B3  depth=2
                 */
                var sheet = package.Workbook.Worksheets.Add("test");
                ExcelRangeBase cellA1 = sheet.Cells["A1"];
                cellA1.Formula = "A2+A3";
                ExcelRangeBase cellA2 = sheet.Cells["A2"];
                cellA2.Formula = "B2";
                ExcelRangeBase cellA3 = sheet.Cells["A3"];
                cellA3.Formula = "B3";

                IEnumerable<IFormulaCellInfo> depthTree = package.Workbook.FormulaParserManager.GetCalculationChainByDepth(cellA1);

                Assert.AreEqual(depthTree.ToList()[0].Address, cellA1.Address);
                Assert.AreEqual(depthTree.ToList()[1].Address, cellA2.Address);
                Assert.AreEqual(depthTree.ToList()[2].Address, cellA3.Address);// B2 and B3 cannot come before A3 in a depth-first traversal
            }
        }


        [TestMethod]
        public void BranchAtDepth0And1()
        {
            using (var package = new ExcelPackage())
            {
                /*
                 *   A1   depth=0
                 *  |   \
                 *  A2   A3   depth=1
                 *  |  \   \ 
                 *  B2  B3  B4    depth=2
                 */
                var sheet = package.Workbook.Worksheets.Add("test");
                ExcelRangeBase cellA1 = sheet.Cells["A1"];
                cellA1.Formula = "A2+A3";
                ExcelRangeBase cellA2 = sheet.Cells["A2"];
                cellA2.Formula = "B2+B3";
                ExcelRangeBase cellA3 = sheet.Cells["A3"];
                cellA3.Formula = "B4";
                

                IEnumerable<IFormulaCellInfo> depthTree = package.Workbook.FormulaParserManager.GetCalculationChainByDepth(cellA1);

                Assert.AreEqual(depthTree.ToList()[0].Address, cellA1.Address);
                Assert.AreEqual(depthTree.ToList()[1].Address, cellA2.Address);
                Assert.AreEqual(depthTree.ToList()[2].Address, cellA3.Address); // B2 and B3 should not come before A3 in a depth-first traversal
            }
        }

        [TestMethod]
        public void BranchAtDepths0And1And2()
        {
            using (var package = new ExcelPackage())
            {
                /*
                 *             A1   depth=0
                 *             |    
                 *         B1  +  B2   depth=1
                 *         |       \     
                 *     C1  +  C2   C3    depth=2
                 *     |      |     \
                 *   D1+D2  D3+D4    D5   depth=3
                 *  
                 */
                var sheet = package.Workbook.Worksheets.Add("test");
                ExcelRangeBase cellA1 = sheet.Cells["A1"];
                cellA1.Formula = "B1+B2";
                ExcelRangeBase cellB1 = sheet.Cells["B1"];
                cellB1.Formula = "C1+C2";
                ExcelRangeBase cellB2 = sheet.Cells["B2"];
                cellB2.Formula = "C3";

                ExcelRangeBase cellC1 = sheet.Cells["C1"];
                cellC1.Formula = "D1+D2";
                ExcelRangeBase cellC2 = sheet.Cells["C2"];
                cellC2.Formula = "D3+D4";
                ExcelRangeBase cellC3 = sheet.Cells["C3"];
                cellC3.Formula = "D5";

                IEnumerable<IFormulaCellInfo> depthTree = package.Workbook.FormulaParserManager.GetCalculationChainByDepth(cellA1);

                Assert.AreEqual(depthTree.ToList()[0].Address, cellA1.Address);
                Assert.AreEqual(depthTree.ToList()[1].Address, cellB1.Address); 
                Assert.AreEqual(depthTree.ToList()[2].Address, cellB2.Address); // C1, C2, and C3 should not come before B2 in a depth-first traversal
                Assert.AreEqual(depthTree.ToList()[3].Address, cellC1.Address); // C1 before C2
                Assert.AreEqual(depthTree.ToList()[4].Address, cellC2.Address); // C2 before C3
                Assert.AreEqual(depthTree.ToList()[5].Address, cellC3.Address); // C3 at last
            }
        }

    }
}
