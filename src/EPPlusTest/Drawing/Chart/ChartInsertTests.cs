using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ChartInsertTests : TestBase
    {
        [TestMethod]
        public void InsertingShouldMoveChartData()
        {
            var fileName = "InsertingShouldMoveChartData.xlsx";

            using (ExcelPackage p = OpenPackage(fileName, true))
            {
                var wb = p.Workbook;
                var ws = wb.Worksheets.Add("Chartinsert");

                ws.Cells["A1"].Value = "CompanyName";
                ws.Cells["B1"].Value = "NumberOfEmployees";
                ws.Cells["C1"].Value = "NumberOfSales";
                ws.Cells["D1"].Value = "DeathToll";

                var companyRange = ws.Cells["A2:A5"];

                int iterationNum = 1;

                for (int i = 2; i < 6; i++)
                {
                    companyRange[i, 1].Value = $"Company{iterationNum}";
                }

                var employeeRange = ws.Cells["B2:B5"];
                employeeRange.Formula = "ROW()*1.5 + COLUMN()";

                var numSalesRange = ws.Cells["C2:C5"];
                numSalesRange.Formula = "ROW() + COLUMN()*2";

                var deathToll = ws.Cells["D2:D5"];
                deathToll.Formula = "ROW()*2 + COLUMN()";

                ws.Calculate();

                var chart = ws.Drawings.AddBarChart("CompanyChart", eBarChartType.ColumnStacked100);

                chart.Series.Add(employeeRange.TakeSingleColumn(0));
                chart.Series.Add(numSalesRange.TakeSingleColumn(0));
                chart.Series.Add(deathToll.TakeSingleColumn(0));

                chart.SetPosition(1, 0, 5, 0);
                chart.SetSize(400, 400);

                SaveAndCleanup(p);
            }

            using (ExcelPackage p = OpenPackage(fileName))
            {
                var ws = p.Workbook.Worksheets.First();

                //foreach (var drawing in ws.Drawings)
                //{
                //    if (drawing.DrawingType == eDrawingType.Chart)
                //    {
                //        var aChart = drawing.As.Chart.Chart;
                //        var chartSerie = drawing.As.Chart.Chart.Series;


                //        //aChart.Series[0].HeaderAddress
                //        //foreach (var serie in chartSerie)
                //        //{
                //        //    foreach (var col in deletedCols)
                //        //    {
                //        //        DeleteColumnFromSeries(ws, serie, col);
                //        //    }
                //        //}
                //    }
                //}

                ws.InsertColumn(1, 1);

                var saveName = GetOutputFile("", $"afterInsert_{fileName}").FullName;
                p.SaveAs(saveName);
            }
        }


        public void DeleteColumnFromSeries(ExcelWorksheet ws, ExcelChartSerie serie, int deletedColumn)
        {
            if (serie.HeaderAddress != null)
            {
                serie.HeaderAddress = ws.Cells[UpdateSerieString(ws, serie.HeaderAddress.Address, deletedColumn)];
            }
            serie.Series = UpdateSerieString(ws, serie.Series, deletedColumn);
            serie.XSeries = UpdateSerieString(ws, serie.XSeries, deletedColumn);
        }

        public string UpdateSerieString(ExcelWorksheet ws, string serieString, int deletedColumn)
        {
            string updatedString = serieString;

            if (!string.IsNullOrEmpty(serieString))
            {
                var newSerieString = DeleteColumnFromAddress(ws, new ExcelAddress(serieString), deletedColumn);

                if (newSerieString != null)
                {
                    updatedString = newSerieString;
                }
            }

            return updatedString;
        }

        public string DeleteColumnFromAddress(ExcelWorksheet ws, ExcelAddressBase address, int deletedColumn)
        {
            if (address != null)
            {
                if (address.Start.Column > deletedColumn)
                {
                    var start = address.Start;
                    var end = address.End;

                    var newAddress = ws.Cells[start.Row, start.Column - 1, end.Row, end.Column - 1];
                    return newAddress.FullAddressAbsolute;
                }
                return address.Address;
            }
            return null;
        }


        [TestMethod]
        public void ColumnCheck()
        {
            using (var p = OpenTemplatePackage("s808_2.xlsx"))
            {
                var ws = p.Workbook.Worksheets["overzicht"];
                List<ExcelRangeColumn> cols = [.. ws.Columns.Where(c => c.Hidden).OrderByDescending(c => c.StartColumn)];

                List<int> deletedCols = new();

                foreach (ExcelRangeColumn col in cols)
                {
                    ws.DeleteColumn(col.StartColumn);
                    deletedCols.Add(col.StartColumn);
                }

                foreach (var drawing in ws.Drawings)
                {
                    if (drawing.DrawingType == eDrawingType.Chart)
                    {
                        var chartSerie = drawing.As.Chart.Chart.Series;

                        foreach (var serie in chartSerie)
                        {
                            foreach (var col in deletedCols)
                            {
                                DeleteColumnFromSeries(ws, serie, col);
                            }
                        }
                    }
                }

                SaveAndCleanup(p);
            }
        }
    }
}
