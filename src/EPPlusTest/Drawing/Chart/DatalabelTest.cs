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
        [ExpectedException(typeof(System.InvalidOperationException))]
        public void ShouldThrowWhenTopPastBottom()
        {
            var sheet3 = _pck.Workbook.Worksheets.Add("LayoutModesTopBottom");

            sheet3.Tables.Add(sheet3.Cells["A1:B1"], "TopBottom");

            var sChart = sheet3.Drawings.AddBarChart("simpleChart", eBarChartType.ColumnStacked);

            sChart.Series.Add(sheet3.Cells["A1"]);
            sChart.Series.Add(sheet3.Cells["B1"]);

            var series = sChart.Series;

            var dataLabel = series[1].DataLabel.DataLabels.Add(0);

            dataLabel.Position = eLabelPosition.Center;

            dataLabel.ShowLegendKey = false;
            dataLabel.ShowValue = true;
            dataLabel.ShowCategory = false;
            dataLabel.ShowSeriesName = false;
            dataLabel.ShowPercent = false;
            dataLabel.ShowBubbleSize = false;

            dataLabel.Layout.ManualLayout.RightMode = eLayoutMode.Edge;
            dataLabel.Layout.ManualLayout.BottomMode = eLayoutMode.Edge;
            dataLabel.Layout.ManualLayout.TopMode = eLayoutMode.Edge;
            dataLabel.Layout.ManualLayout.LeftMode = eLayoutMode.Edge;

            dataLabel.Layout.ManualLayout.Left = 30;
            dataLabel.Layout.ManualLayout.Top = 25;

            dataLabel.Layout.ManualLayout.Right = 40;
            dataLabel.Layout.ManualLayout.Bottom = 20;
        }

        [TestMethod]
        [ExpectedException(typeof(System.InvalidOperationException))]
        public void ShouldThrowWhenRightMoreLeftThanLeft()
        {
            var sheet3 = _pck.Workbook.Worksheets.Add("LayoutModesLeftRight");

            sheet3.Tables.Add(sheet3.Cells["A1:B1"], "LeftRight");

            var sChart = sheet3.Drawings.AddBarChart("simpleChart", eBarChartType.ColumnStacked);

            sChart.Series.Add(sheet3.Cells["A1"]);
            sChart.Series.Add(sheet3.Cells["B1"]);

            var series = sChart.Series;

            var dataLabel = series[1].DataLabel.DataLabels.Add(0);

            dataLabel.Position = eLabelPosition.Center;

            dataLabel.ShowLegendKey = false;
            dataLabel.ShowValue = true;
            dataLabel.ShowCategory = false;
            dataLabel.ShowSeriesName = false;
            dataLabel.ShowPercent = false;
            dataLabel.ShowBubbleSize = false;

            dataLabel.Layout.ManualLayout.RightMode = eLayoutMode.Edge;
            dataLabel.Layout.ManualLayout.BottomMode = eLayoutMode.Edge;
            dataLabel.Layout.ManualLayout.TopMode = eLayoutMode.Edge;
            dataLabel.Layout.ManualLayout.LeftMode = eLayoutMode.Edge;

            dataLabel.Layout.ManualLayout.Left = 30;
            dataLabel.Layout.ManualLayout.Top = 25;

            dataLabel.Layout.ManualLayout.Right = 20;
            dataLabel.Layout.ManualLayout.Bottom = 30;
        }

        [TestMethod]
        public void EdgeTest()
        {
            var sheet3 = _pck.Workbook.Worksheets.Add("LayoutModeEdge");

            sheet3.Tables.Add(sheet3.Cells["A1:B1"], "table1");

            sheet3.Cells["A1"].Value = 5;
            sheet3.Cells["B1"].Value = 10;

            var sChart = sheet3.Drawings.AddBarChart("simpleChart", eBarChartType.ColumnStacked);

            sChart.Series.Add(sheet3.Cells["A1"]);
            sChart.Series.Add(sheet3.Cells["B1"]);

            var series = sChart.Series;

            var dataLabel = series[1].DataLabel.DataLabels.Add(0);

            dataLabel.Position = eLabelPosition.Center;

            dataLabel.ShowLegendKey = false;
            dataLabel.ShowValue = true;
            dataLabel.ShowCategory = false;
            dataLabel.ShowSeriesName = false;
            dataLabel.ShowPercent = false;
            dataLabel.ShowBubbleSize = false;

            //dataLabel.Layout.ManualLayout.WidthMode = eLayoutMode.Edge;
            //dataLabel.Layout.ManualLayout.HeightMode = eLayoutMode.Edge;
            //dataLabel.Layout.ManualLayout.TopMode = eLayoutMode.Edge;
            //dataLabel.Layout.ManualLayout.LeftMode = eLayoutMode.Edge;

            var manualLayout = dataLabel.Layout.ManualLayout;

            manualLayout.LeftMode = eLayoutMode.Edge;
            manualLayout.TopMode = eLayoutMode.Edge;
            manualLayout.WidthMode = eLayoutMode.Edge;
            manualLayout.HeightMode = eLayoutMode.Edge;

            manualLayout.Top = 20;
            manualLayout.Left = 20;

            manualLayout.Width = 15;
            manualLayout.Height = 15;

            ////manualLayout.Left = 45;
            //manualLayout.Width = 10;

            ////manualLayout.Top = 45;
            //manualLayout.Height = 10;
            //dataLabel.Layout.ManualLayout.Top = 25;

            //dataLabel.Layout.ManualLayout.Width = 20;
            //dataLabel.Layout.ManualLayout.Height = 30;
        }

        [TestMethod]
        public void EdgeTest2()
        {
            var sheet3 = _pck.Workbook.Worksheets.Add("LayoutModeFactor");

            sheet3.Tables.Add(sheet3.Cells["A1:B1"], "modeFactorTable");

            sheet3.Cells["A1"].Value = 5;
            sheet3.Cells["B1"].Value = 10;

            var sChart = sheet3.Drawings.AddBarChart("simpleChart", eBarChartType.ColumnStacked);

            sChart.Series.Add(sheet3.Cells["A1"]);
            sChart.Series.Add(sheet3.Cells["B1"]);

            var series = sChart.Series;

            var dataLabel = series[1].DataLabel.DataLabels.Add(0);

            dataLabel.Position = eLabelPosition.Center;

            dataLabel.ShowLegendKey = false;
            dataLabel.ShowValue = true;
            dataLabel.ShowCategory = false;
            dataLabel.ShowSeriesName = false;
            dataLabel.ShowPercent = false;
            dataLabel.ShowBubbleSize = false;

            //dataLabel.Layout.ManualLayout.WidthMode = eLayoutMode.Edge;
            //dataLabel.Layout.ManualLayout.HeightMode = eLayoutMode.Edge;
            //dataLabel.Layout.ManualLayout.TopMode = eLayoutMode.Edge;
            //dataLabel.Layout.ManualLayout.LeftMode = eLayoutMode.Edge;

            //dataLabel.Layout.ManualLayout.LeftMode = eLayoutMode.Edge;
            //dataLabel.Layout.ManualLayout.WidthMode = eLayoutMode.Edge;


            //dataLabel.Layout.ManualLayout.TopMode = eLayoutMode.Edge;
            //dataLabel.Layout.ManualLayout.HeightMode = eLayoutMode.Edge;

            dataLabel.Layout.ManualLayout.Left = 20;
            dataLabel.Layout.ManualLayout.Width = 10;
            dataLabel.Layout.ManualLayout.Height = 20;


            //dataLabel.Layout.ManualLayout.Top = 45;
            //dataLabel.Layout.ManualLayout.Height = 55;
            //dataLabel.Layout.ManualLayout.Top = 25;

            //dataLabel.Layout.ManualLayout.Width = 20;
            //dataLabel.Layout.ManualLayout.Height = 30;
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
