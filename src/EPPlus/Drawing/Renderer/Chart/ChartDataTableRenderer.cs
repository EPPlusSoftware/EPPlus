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
using EPPlus.Export.ImageRenderer.Svg.Chart.Util;
using EPPlusImageRenderer;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.String;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Drawing.Renderer.Chart
{
    internal class ChartDataTableRenderer : ChartDrawingObject, ILegendKeyContainer
    {
        ExcelChartDataTable _dataTable;
        float _marginItemsWidth;
        public double _maxWidth, _maxHeight;
        public float MarginItemsWidth => _marginItemsWidth;
        List<TextMeasurement> _seriesHeadersMeasure = new List<TextMeasurement>();
        public double MaxWidth => _maxWidth;

        public double MaxHeight => _maxHeight;
        public List<TextMeasurement> SeriesHeadersMeasure => _seriesHeadersMeasure;

        public List<DrawingLegendSerie> LegendColumn = new List<DrawingLegendSerie>();
        public List<DrawingLegendSerie> SeriesIcon { get;  }=new List<DrawingLegendSerie>();
        public List<List<DrawingTextBody>> DataTableRenderItems { get; set; } = new List<List<DrawingTextBody>>();
        double _dataTableWidth, _columnsWidth;
        internal ChartDataTableRenderer(ChartRenderer svgChart) : base(svgChart)
        {
            _dataTable = svgChart.Chart.PlotArea.DataTable;
            Rectangle = new RectRenderItem(svgChart.Plotarea.Rectangle.Bounds);

            MeasurementFont mf;
            if(_dataTable.HasFont)
            {
                mf = _dataTable.Font.GetMeasureFont();
            }
            else
            {
                mf = Chart.Font.GetMeasureFont();
            }

            _marginItemsWidth = mf.Size / 2;
            _maxWidth = svgChart.Plotarea.GetPlotAreaWidth(Rectangle);
            _maxHeight = svgChart.Plotarea.GetPlotAreaHeight(Rectangle);
            var items = new List<List<DrawingTextBox>>();
            var headers = new List<DrawingTextBox>();
            items.Add(headers);
            
            var tm = svgChart.TextMeasurer;
            double entryWidth = 0, entryHeight=0;
            var values = svgChart.HorizontalAxis.Axis.GetAxisValues(out _, out _, out _);
            var horizontaValues = GetFormattedValues(svgChart, values);
            foreach (var v in horizontaValues)
            {
                var tb = new DrawingTextBox(svgChart.Chart, Rectangle.Bounds, MaxWidth, MaxHeight);
                tb.AddText(v);
                headers.Add(tb);
                var size = tm.MeasureText(v, mf);
                if (entryWidth<size.Width)
                {
                    entryWidth = size.Width;
                }
                if(entryHeight<size.Height)
                {
                    entryHeight = size.Height;
                }
                _seriesHeadersMeasure.Add(size);
            }

            ExcelDrawingParagraph paragraph;
            if (_dataTable.HasFont)
            {
                paragraph = _dataTable.TextBody.Paragraphs.FirstOrDefault();
            }
            else
            {
                paragraph = null;
            }

            _dataTableWidth = GetDataTableWidth(svgChart, _dataTable);
            _columnsWidth = (_dataTableWidth - entryWidth) / headers.Count;

            //Create the source data for the data table from the series.
            DrawingLegendSerie pSls = null;
            int index = 0;
            int seriesIndex=0;
            var maxIconLength = LegendIconRenderer.GetIconLength(Chart, entryHeight);
            foreach (var ct in svgChart.Chart.PlotArea.ChartTypes)
            {
                foreach (ExcelChartStandardSerie serie in ct.Series)
                {
                    //Create the legend column
                    var sls = new DrawingLegendSerie();
                    if(ct.IsTypeLine())
                    {
                        LegendIconRenderer.SetLineLegend(ChartRenderer, this, ct, index, pSls, serie, sls, entryWidth, entryHeight, maxIconLength);
                    }
                    else if(ct.IsTypePie())
                    {
                        LegendIconRenderer.SetPieLegend(ChartRenderer, this, ct, index, pSls, serie, sls, entryWidth, entryHeight, maxIconLength);
                    }
                    else if(ct.IsTypeColumn() || ct.IsTypeBar())
                    {
                        LegendIconRenderer.SetBarLegend(ChartRenderer, this, ct, index, pSls, serie, sls, entryWidth, entryHeight, maxIconLength);
                    }
                    var rows = AddSerieValues(svgChart, _maxWidth, _maxHeight, serie);
                    LegendColumn.Add(sls);
                    pSls = sls;
                    var formattedValues = serie.GetValues(false, true);
                    var l=new List<DrawingTextBody>();
                    foreach (var v in formattedValues)
                    {
                        var tb = new DrawingTextBody(RenderContext, ChartRenderer.Chart, Rectangle.Bounds, true);
                        if (paragraph == null)
                        {
                            tb.AddParagraph(v.ToString());
                        }
                        else
                        {
                            tb.ImportParagraph(paragraph, 0, v.ToString());
                        }
                        l.Add(tb);
                    }
                    DataTableRenderItems.Add(l);
                }
                seriesIndex++;
            }

            SetColumnWidth(entryHeight);
        }


        private void SetColumnWidth(double entryHeight)
        {
            var ledgedColWidth = GetLegendColWidth();
            foreach(var lc in LegendColumn)
            {
                if (_dataTable.ShowKeys) 
                {
                    lc.MarkerIcon.Bounds.Left += ledgedColWidth;
                    lc.MarkerBackground.Bounds.Left += ledgedColWidth;
                    lc.SeriesIcon.Bounds.Left += ledgedColWidth;
                    lc.Textbox.Bounds.Left += ledgedColWidth;
                }
                else
                {
                    lc.Textbox.Left = ledgedColWidth;
                }
                ledgedColWidth += _columnsWidth;
            }

            double y = entryHeight;
            foreach (var row in DataTableRenderItems)
            {
                var x = ledgedColWidth;
                foreach (var cell in row)
                {
                    cell.Bounds.Left = x;
                    cell.Bounds.Top = y;
                    x += _columnsWidth;
                }
                y += entryHeight;
           }

        }

        const float MarginIconText = 1.5f;
        private double GetLegendColWidth()
        {
            if (LegendColumn.Count == 0) return 0;
            var w = LegendColumn.Max(x=>x.Textbox.Width);
            if(_dataTable.ShowKeys)
            {
                return w + LegendColumn.First().MarkerIcon.Bounds.Width + MarginIconText;
            }
            return w;
        }

        private double GetDataTableWidth(ChartRenderer svgChart, ExcelChartDataTable chartDataTable)
        {
            var width = svgChart.ChartArea.Rectangle.Width - svgChart.ChartArea.LeftMargin - svgChart.ChartArea.RightMargin;
            if(svgChart.Legend==null || svgChart.Chart.Legend.Position==eLegendPosition.Top || svgChart.Chart.Legend.Position == eLegendPosition.Bottom)
            {
                return width;
            }
            else
            {
                return width - svgChart.Legend.Rectangle.Width - svgChart.Legend.RightMargin;
            }
        }

        private List<string> GetFormattedValues(ChartRenderer svgChart, List<object> values)
        {
            var format = svgChart.HorizontalAxis.Axis.FormatOrFirstValueFormat;
            var nf = new ExcelFormatTranslator(format, 0);
            //Excel replaces the format with a default date format if the axis is date based.
            if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
            {
                if (format == "m/d/yyyy")
                {
                    var sdFormat = ExcelNumberFormat.GetFromBuildInFromID(14); //14 is standard regional short date.
                    nf = new ExcelFormatTranslator(sdFormat, 14);
                }
            }
            var displayValues = new List<string>();
            foreach (var v in values)
            {
                var s = ValueToTextHandler.FormatValue(v, false, nf, null, out bool isValidFormat);
                displayValues.Add(s);
            }
            return displayValues;
        }

        private RenderItem GetIcon(ChartRenderer svgChart, ExcelChart ct, ExcelChartStandardSerie serie, DrawingLegendSerie pSls, int serieIndex, int index, double entryWidth, double entryHeight)
        {
            if(ct.IsTypeLine())
            {
                return LegendIconRenderer.GetLineSeriesIcon(svgChart, this, serie, pSls, entryWidth, entryHeight);
            }
            else if(ct.IsTypeBar() || ct.IsTypeColumn())
            {
                return LegendIconRenderer.GetBarSeriesIcon(svgChart, ct, this, (ExcelBarChartSerie)serie, pSls, entryWidth, entryHeight, serieIndex, index);
            }
            else
            {
                return LegendIconRenderer.GetPieSeriesIcon(svgChart, ct, this, (ExcelPieChartSerie)serie, pSls, entryWidth, entryHeight, index);
            }
        }

        private List<DrawingTextBox> AddSerieValues(ChartRenderer svgChart, double maxWidth, double maxHeight, ExcelChartSerie serie)
        {
            var ret = new List<DrawingTextBox>();
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
                        ret.Add(tb);
                    }
                }
            }
            else if (serie.StringLiteralsY != null && serie.StringLiteralsY.Length > 0)
            {
                foreach (var se in serie.StringLiteralsY)
                {
                    var tb = new DrawingTextBox(svgChart.Chart, Rectangle.Bounds, maxWidth, maxHeight);
                    tb.AddText(se);
                    ret.Add(tb);
                }
            }
            else if (serie.NumberLiteralsY != null && serie.NumberLiteralsY.Length > 0)
            {
                foreach (var nl in serie.NumberLiteralsY)
                {
                    var tb = new DrawingTextBox(svgChart.Chart, Rectangle.Bounds, maxWidth, maxHeight);
                    tb.AddText(nl.ToString());
                    ret.Add(tb);
                }
            }

            return ret;
        }

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            //foreach(var tb in )
        }        
    }
}