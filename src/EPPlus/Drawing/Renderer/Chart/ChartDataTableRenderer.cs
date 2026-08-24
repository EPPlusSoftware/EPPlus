/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  20/08/2026         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/

using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.RenderItems.SvgItem;
using EPPlus.DrawingRenderer.ShapeDefinitions;
using EPPlusImageRenderer;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Drawing.Renderer.Chart
{
    internal class ChartDataTableRenderer : ChartDrawingObject
    {
        internal ChartDataTableRenderer(ChartRenderer svgChart) : base(svgChart)
        {
            var chartDataTable = svgChart.Chart.PlotArea.DataTable;
            Rectangle = new RectRenderItem(svgChart.Plotarea.Rectangle.Bounds);
            var maxWidth = svgChart.Plotarea.GetPlotAreaWidth(Rectangle);
            var maxHeight = svgChart.Plotarea.GetPlotAreaHeight(Rectangle);
            var items = new List<List<DrawingTextBox>>();
            var headers = new List<DrawingTextBox>();
            items.Add(headers);
            foreach (var v in svgChart.HorizontalAxis.AxisValues)
            {
                var tb = new DrawingTextBox(svgChart.Chart, Rectangle.Bounds, maxWidth, maxHeight);
                tb.AddText(v);
                headers.Add(tb);
            }            
            //Create the source data for the data table from the series.
            foreach (var ct in svgChart.Chart.PlotArea.ChartTypes)
            {
                foreach (var serie in ct.Series)
                {
                    bool flowControl = AddSerieValues(svgChart, maxWidth, maxHeight, serie);
                    if (!flowControl)
                    {
                        return;
                    }
                }
            }
        }

        private bool AddSerieValues(ChartRenderer svgChart, double maxWidth, double maxHeight, Drawing.Chart.ExcelChartSerie serie)
        {
            if (string.IsNullOrEmpty(serie.Series))
            {
                var a = new ExcelAddressBase(serie.Series);
                var ws = svgChart.Chart.WorkSheet;
                if (ws != null)
                {
                    var range = ws.Cells[a.Address];
                    foreach (var cell in range)
                    {
                        var tb = new DrawingTextBox(svgChart.Chart, Rectangle.Bounds, maxWidth, maxHeight);
                        tb.AddText(cell.Text);
                    }
                }
            }
            else if (serie.StringLiteralsY != null && serie.StringLiteralsY.Length > 0)
            {
                foreach (var se in serie.StringLiteralsY)
                {
                    var tb = new DrawingTextBox(svgChart.Chart, Rectangle.Bounds, maxWidth, maxHeight);
                    tb.AddText(se);
                }
            }
            else if (serie.NumberLiteralsY != null && serie.NumberLiteralsY.Length > 0)
            {
                foreach (var nl in serie.NumberLiteralsY)
                {
                    var tb = new DrawingTextBox(svgChart.Chart, Rectangle.Bounds, maxWidth, maxHeight);
                    tb.AddText(nl.ToString());
                }
            }
            else
            {
                return false;
            }

            return true;
        }

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            base.AppendRenderItems(renderItems);
        }        
    }
}