using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class DatalabelTest : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet dSheet;
        static ExcelWorksheet cSheet;

        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("DataLabel.xlsx", true);
            dSheet = _pck.Workbook.Worksheets.Add("DataSheet");
            cSheet = _pck.Workbook.Worksheets.Add("ChartSheet");

            var range = new ExcelRange(dSheet, "A1:G10");
            var table = dSheet.Tables.Add(range, "DataTable");
            table.ShowHeader = true;
            dSheet.Cells["A2:G10"].Formula = "ROW() + COLUMN()";
            dSheet.Cells["D2:D10"].Value = 0.1;
            dSheet.Cells["E2:E10"].Value = 1;
            dSheet.Cells["F2:F10"].Value = 0.5;


            var pTable = cSheet.PivotTables.Add(cSheet.Cells["A1"], range, "NewPivotTable");

            pTable.DataFields.Add(pTable.Fields["Column1"]);
            pTable.DataFields.Add(pTable.Fields["Column2"]);
            pTable.DataFields.Add(pTable.Fields["Column3"]);
            pTable.DataFields.Add(pTable.Fields["Column4"]);
            pTable.DataFields.Add(pTable.Fields["Column5"]);
            pTable.DataFields.Add(pTable.Fields["Column6"]);
            pTable.DataFields.Add(pTable.Fields["Column7"]);

            pTable.ShowColumnHeaders = true;
            pTable.DataOnRows = false;

            var bChart = cSheet.Drawings.AddBarChart("PivotChartTestTwo", eBarChartType.ColumnStacked, pTable);

            bChart.DataLabel.ShowValue = true;
            bChart.DataLabel.ShowLeaderLines = true;
        }

        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }

        // s679
        //Manual layout when labels are atop eachother
        [TestMethod]
        public void AddingManualLayout()
        {
            var bChart = cSheet.Drawings[0].As.Chart.BarChart;

            bChart.XAxis.RemoveGridlines(true, true);
            bChart.YAxis.RemoveGridlines(true, true);

            for (int i = 0; i < bChart.Series.Count; i++)
            {
                var label = bChart.Series[i].DataLabel.DataLabels.Add(0);
                AdjustDataLabelItem(ref label);
                if (i == 3 || i == 2)
                {
                    label.Layout.ManualLayout.Top = 5;
                }
                else if (i == 4)
                {
                    label.Layout.ManualLayout.Top = 0;

                }
                else if (i == 5)
                {
                    label.Layout.ManualLayout.Top = -5;
                }
            }
        }

        [TestMethod]
        public void simpleChartTest()
        {
            //using (var package = OpenTemplatePackage("SimpleChart.xlsx"))
            //{
            var bChart = cSheet.Drawings[0].As.Chart.BarChart;

            _pck.Workbook.Worksheets.Add("LayoutModes");

            var chart2 = cSheet.Drawings["Chart 3"];

            var barChart = chart2.As.Chart.BarChart;

            var series = barChart.Series;

            var dataLabel = series[1].DataLabel.DataLabels.Add(0);

            dataLabel.Position = eLabelPosition.Center;

            dataLabel.ShowLegendKey = false;
            dataLabel.ShowValue = true;
            dataLabel.ShowCategory = false;
            dataLabel.ShowSeriesName = false;
            dataLabel.ShowPercent = false;
            dataLabel.ShowBubbleSize = false;

            dataLabel.Layout.ManualLayout.WidthMode = eLayoutMode.Edge;
            dataLabel.Layout.ManualLayout.HeightMode = eLayoutMode.Edge;
            dataLabel.Layout.ManualLayout.TopMode = eLayoutMode.Edge;
            dataLabel.Layout.ManualLayout.LeftMode = eLayoutMode.Edge;

            dataLabel.Layout.ManualLayout.Left = 30;
            dataLabel.Layout.ManualLayout.Top = 25;

            dataLabel.Layout.ManualLayout.Width = 40;
            dataLabel.Layout.ManualLayout.Height = 20;

            //dataLabel.Layout.ManualLayout.Width = 5d;
            //dataLabel.Layout.ManualLayout.Height = 5d;

            //dataLabel.Layout.ManualLayout.Height = -10d;

            SaveAndCleanup(package);
        //}
        }


        void AdjustDataLabelItem(ref ExcelChartDataLabelItem label)
        {
            label.ShowSeriesName = false;
            label.ShowCategory = false;
            label.ShowLegendKey = false;
            label.ShowLeaderLines = true;
            label.ShowValue = true;

            label.Position = eLabelPosition.Center;

            label.Layout.ManualLayout.Left = -30;
        }
    }
}
