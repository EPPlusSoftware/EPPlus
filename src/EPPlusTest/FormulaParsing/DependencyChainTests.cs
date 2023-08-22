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
        /// <summary>
        /// Verifies that the breadth-first chain for this tree is A1 
        /// </summary>
        [TestMethod]
        public void ChainOrder_OneLink()
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

        /// <summary>
        /// Verifies that the breadth-first chain for this tree is A1 > A2 
        /// </summary>
        [TestMethod]
        public void ChainOrder_TwoLinks()
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

        /// <summary>
        /// Verifies that the breadth-first chain for this tree is A1 > A2 > A3
        /// 
        ///  A1   depth=0
        /// |  \
        /// A2  A3  depth=1
        /// |   |
        /// B2  B3  depth=2
        /// </summary>
        [TestMethod]
        public void ChainOrder_BranchAtOneDepth()
        {
            using (var package = new ExcelPackage())
            {
                ///
                ///  A1   depth=0
                /// |  \
                /// A2  A3  depth=1
                /// |   |
                /// B2  B3  depth=2
                ///
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

        /// <summary>
        /// Verifies that the breadth-first chain for this tree is A1 > A2 > A3
        /// 
        ///  A1   depth=0
        /// |   \
        /// A2   A3   depth=1
        /// |  \   \ 
        /// B2  B3  B4    depth=2
        /// </summary>
        [TestMethod]
        public void ChainOrder_BranchAtTwoDepths()
        {
            using (var package = new ExcelPackage())
            {               
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

        /// <summary>
        /// Verifies that the breadth-first chain for this tree is A1 > B1 > B2 > C1 > C2 > C3
        /// 
        ///            A1   depth=0
        ///            |    
        ///        B1  +  B2   depth=1
        ///        |       \     
        ///    C1  +  C2   C3    depth=2
        ///    |      |     \
        ///  D1+D2  D3+D4    D5   depth=3
        ///
        /// </summary>
        [TestMethod]
        public void ChainOrder_BranchAtThreeDepths()
        {
            using (var package = new ExcelPackage())
            {                
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


        /// <summary>
        /// This test fails because breadth-first traversal is being simulated in a depth-first algorithm
        /// Despite A1 referencing B2 "higher" than C1's reference to B2, C1's reference is encountered first, so B2 is treated as being "lower" than C1
        /// Only truly breadth-first traversal can produce the true breadth-first ordered results
        /// </summary>
        [TestMethod]
        public void ChainOrder_MultipleReferences()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                ExcelRangeBase A1 = sheet.Cells["A1"];
                A1.Formula = "B1+B2";  //<--- this instance of a reference B2 will be found second during depth-first tree traversal
                ExcelRangeBase B1 = sheet.Cells["B1"];
                B1.Formula = "C1";
                ExcelRangeBase B2 = sheet.Cells["B2"];
                B2.Formula = "C2";
                ExcelRangeBase C1 = sheet.Cells["C1"];
                C1.Formula = "B2"; //<--- this instance of a reference B2 will be found first during depth-first tree traversal

                IEnumerable<IFormulaCellInfo> depthTree = package.Workbook.FormulaParserManager.GetCalculationChainByDepth(A1);
                Assert.AreEqual(depthTree.ToList()[0].Address, A1.Address);
                Assert.AreEqual(depthTree.ToList()[1].Address, B1.Address);
                Assert.AreEqual(depthTree.ToList()[2].Address, B2.Address); // B2 should come before C1 (but this is a result of simulating breadth-first traversal from a depth-first traversal)
                Assert.AreEqual(depthTree.ToList()[3].Address, C1.Address); // C1 should be after B2
            }
        }


        /// <summary>
        /// This is a template for testing dependency chains from an excel doc
        /// </summary>
        [TestMethod]
        public void ChainOrder_FromFile()
        {
            using (var package = new ExcelPackage(@"Dependency Chain Depth.xlsm"))
            {
                var sheet = package.Workbook.Worksheets["test"];
                ExcelRangeBase A1 = sheet.Cells["A1"];
                IEnumerable<IFormulaCellInfo> depthTree = package.Workbook.FormulaParserManager.GetCalculationChainByDepth(A1);
            }
        }



    }
}
