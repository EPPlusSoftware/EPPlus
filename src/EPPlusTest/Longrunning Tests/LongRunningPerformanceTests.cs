using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Longrunning_Tests
{
    [TestClass, Ignore("Remove this attribute when needed, should not run in CI/CD.")]
    public class LongRunningPerformanceTests : TestBase
    {
        //i2084
        [TestMethod]
        public void s912_Alternate()
        {
            //Optimizing for not overwriting existing styles
            using (var package = OpenPackage("s912_alt.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("F1");

                int nbLines = 10000;
                int nbCols = 100;

                var sw = new Stopwatch();
                sw.Start();

                for (int i = 1; i <= nbLines; i++)
                {
                    for (int j = 1; j <= nbCols; j++)
                    {
                        var cell = sheet.Cells[i, j];
                        var cellNumberFormat = cell.Style.Numberformat;
                        cell.Value = 123;
                    }
                }
                sw.Stop();

                var seconds = sw.Elapsed.TotalSeconds;
                Assert.IsTrue(seconds < 10.0D);

                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void s912()
        {
            using (var package = OpenPackage("s912.xlsx", true))
            {
                var sheet = package.Workbook.Worksheets.Add("F1");

                int nbLines = 10000;
                int nbCols = 100;

                var sw = new Stopwatch();
                sw.Start();

                // Uncommenting one of these lines changes the performance of the for loops.
                // At the end of each line is the measured time of the whole program, when this
                // specific line is uncommented. When no line is uncommented, the measured time
                // is 12.7s.
                //
                // sheet.Cells[1, 1, nbLines, nbCols].Style.Numberformat.Format = "#"; // 7.4s
                // sheet.Cells[1, 1, nbLines, nbCols].Style.Locked = true; // 7.4s
                //sheet.Cells[1, 1, nbLines, nbCols].Value = 1; // 19.5s // uncommenting this alone is ~2s after fix. With below about 3s. 
                // sheet.Cells[1, 1, nbLines, nbCols].Value = ""; // 18s
                // sheet.InsertColumn(1, nbCols); // 12.9
                // sheet.InsertColumn(1, nbCols, 1); // 7.8s

                for (int i = 1; i <= nbLines; i++)
                {
                    for (int j = 1; j <= nbCols; j++)
                    {
                        sheet.SetValue(i, j, 123);
                    }
                }

                var seconds = sw.Elapsed.TotalSeconds;
                sw.Stop();

                //seconds was ~1.5-1.7 locally in 8.0.9
                //Made test check if over 10 seconds in case of slow appveyor
                Assert.IsTrue(seconds < 10.0D);

                SaveAndCleanup(package);
            }
        }
    }
}
