using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class CancelCalculationTests
    {
        const int WaitTimeMs = 150;

        [TestMethod]
        public void CancelCalculation()
        {
            using var package = CreateHeavyChain(chainLength: 1500, sheetCount: 3);
            using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(WaitTimeMs));

            var sw = Stopwatch.StartNew();
            try
            {
                package.Workbook.Calculate(opt =>
                {
                    opt.CancellationToken = cts.Token;
                });
                Assert.Fail("Expected OperationCanceledException was not thrown.");
            }
            catch (OperationCanceledException)
            {
                sw.Stop();
                Debug.WriteLine($"Calculation cancelled after {sw.Elapsed.TotalSeconds:F2} seconds.");
                Assert.IsTrue(sw.Elapsed.TotalSeconds < 10, "Cancellation took too long.");
                Assert.IsTrue(package.Workbook.IsCalculationInconsistent);
            }
        }

        [TestMethod]
        public void Save_AfterCancelledCalculation_ThrowsInvalidOperationException()
        {
            using var package = CreateHeavyChain(chainLength: 1500, sheetCount: 3);
            using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(WaitTimeMs));
            using var outputStream = new MemoryStream();

            try
            {
                package.Workbook.Calculate(opt =>
                {
                    opt.CancellationToken = cts.Token;
                });
            }
            catch (OperationCanceledException) { /* expected */ }

            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                package.SaveAs(outputStream);
            });
        }

        [TestMethod]
        public void CancelCalculation_AlreadyCancelledToken_ThrowsImmediately()
        {
            using var package = CreateHeavyChain(chainLength: 1500, sheetCount: 3);
            using var cts = new CancellationTokenSource();
            cts.Cancel(); // Signal before calculate

            Assert.ThrowsExactly<OperationCanceledException>(() =>
            {
                package.Workbook.Calculate(opt => opt.CancellationToken = cts.Token);
            });
            Assert.IsTrue(package.Workbook.IsCalculationInconsistent);
        }

        [TestMethod]
        public void CancelCalculation_RecalculatePoisonedWorkbook_ThrowsInvalidOperationException()
        {
            using var package = CreateHeavyChain(chainLength: 1500, sheetCount: 3);
            using var cts = new CancellationTokenSource(TimeSpan.FromMilliseconds(WaitTimeMs));

            try
            {
                package.Workbook.Calculate(opt => opt.CancellationToken = cts.Token);
            }
            catch (OperationCanceledException) { /* expected */ }

            Assert.ThrowsExactly<InvalidOperationException>(() =>
            {
                package.Workbook.Calculate(); // Should throw — workbook is poisoned
            });
        }

        [TestMethod]
        public void CancelCalculation_FromAnotherThread()
        {
            using var package = CreateHeavyChain(chainLength: 1500, sheetCount: 3);
            using var cts = new CancellationTokenSource();
            Exception caughtException = null;

            var calcThread = new Thread(() =>
            {
                try
                {
                    package.Workbook.Calculate(opt => opt.CancellationToken = cts.Token);
                }
                catch (OperationCanceledException ex)
                {
                    caughtException = ex;
                }
            });

            calcThread.Start();
            Thread.Sleep(WaitTimeMs); // Let calculation run for a while
            cts.Cancel();             // Cancel from this (main) thread
            calcThread.Join(TimeSpan.FromSeconds(10)); // Wait for calc thread to finish

            Assert.IsFalse(calcThread.IsAlive, "Calculation thread did not terminate.");
            Assert.IsInstanceOfType<OperationCanceledException>(caughtException);
            Assert.IsTrue(package.Workbook.IsCalculationInconsistent);
        }



        public static ExcelPackage CreateHeavyChain(int chainLength = 1_000, int sheetCount = 3)
        {
            var package = new ExcelPackage();

            for (int s = 1; s <= sheetCount; s++)
            {
                var ws = package.Workbook.Worksheets.Add($"Sheet{s}");
                ws.Cells[1, 1].Value = 1;

                for (int row = 2; row <= chainLength; row++)
                {
                    // SUMPRODUCT over a growing range — O(N²) total work
                    ws.Cells[row, 1].Formula = $"SUMPRODUCT(A$1:A{row - 1})+1";
                }
            }

            return package;
        }
    }
}
