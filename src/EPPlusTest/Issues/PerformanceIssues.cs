using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.Diagnostics;

namespace EPPlusTest.Issues
{
    [TestClass]
    public class PerformanceIssues : TestBase
    {
        [TestMethod]
        public void s912()
        {

            using (var package = OpenPackage("s912Output.xlsx"))
            {
                var sheet = package.Workbook.Worksheets.Add("F1");

                int nbLines = 10000;
                int nbCols = 100;

                var sw = Stopwatch.StartNew();

                sheet.Cells[1, 1].Style.Numberformat.Format = "#";
                //var someName = sheet.Cells[1, 1, nbLines, nbCols].StyleName;

                //sheet.Cells[1, 1, nbLines, nbCols].StyleName = package.Workbook.Styles.NamedStyles[0].Name;
                //sheet.Calculate();
                //foreach (var col in cols)
                //{
                //    col.Style.Indent = 0;
                //}

                // Uncommenting one of these lines changes the performance of the for loops.
                // At the end of each line is the measured time of the whole program, when this
                // specific line is uncommented. When no line is uncommented, the measured time
                // is 12.7s.
                //

                //ExcelWorksheet.GetStyleID
                //
                //object[,] arr = { 123d };
                //sheet.SetRangeValueInner(1, 1, nbLines, nbCols, arr, false);
                //sheet.Cells[1, 1, nbLines, nbCols].Value = 123;
                //sheet.Cells[1, 1, nbLines, nbCols].Style.Numberformat.Format = "#"; // 7.4s
                //sheet.Cells[1, 1, nbLines, nbCols].Style.Locked = true; // 7.4s
                //sheet.Cells[1, 1, nbLines, nbCols].setd
                sheet.Cells[1, 1, nbLines, nbCols].Value = 1; // 19.5s
                //sheet.Cells[1, 1, nbLines, nbCols].Value = ""; // 18s
                //// sheet.InsertColumn(1, nbCols); // 12.9
                //// sheet.InsertColumn(1, nbCols, 1); // 7.8s

                //var measuredTime1 = sw.ElapsedMilliseconds;
                //sw.Stop();

                ////sw.Start();
                //for (int i = 1; i <= nbLines; i++)
                //{
                //    for (int j = 1; j <= nbCols; j++)
                //    {
                //        sheet.SetValue(i, j, 123);
                //        //sheet.Cells[i,j].Value = 123;

                //    }
                //}

                var measuredTime = sw.ElapsedMilliseconds;

                sw.Stop();
                SaveAndCleanup(package);
            }
        }

        [TestMethod]
        public void EnsurePerformanceOfGetStyleID()
        {
            var package = new ExcelPackage();

            var sheet = package.Workbook.Worksheets.Add("F1");

            var sw = Stopwatch.StartNew();
            sheet.GetStyleId(3, 1000);

            var measuredTime = sw.ElapsedMilliseconds;
            sw.Stop();

        }
    }
}
