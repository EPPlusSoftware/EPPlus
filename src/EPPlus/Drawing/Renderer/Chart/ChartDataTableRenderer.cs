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
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Core;
using OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Renderer.Chart.Defaults;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.String;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace OfficeOpenXml.Drawing.Renderer.Chart
{
    internal class ChartDataTableRenderer : ChartDrawingObjectWithBackground, ILegendKeyContainer
    {
        ExcelChartDataTable _dataTable;
        float _marginItemsWidth;
        public double _maxWidth, _maxHeight;
        public float MarginItemsWidth => _marginItemsWidth;
        List<TextMeasurement> _seriesHeadersMeasure = new List<TextMeasurement>();
        public double MaxWidth => _maxWidth;

        public double MaxHeight => _maxHeight;
        public List<TextMeasurement> SeriesHeadersMeasure => _seriesHeadersMeasure;

        public List<DrawingLegendSerie> SeriesIcon { get;  }=new List<DrawingLegendSerie>();
        public List<List<DrawingTextBody>> DataTableRenderItems { get; set; } = new List<List<DrawingTextBody>>();

        internal override Color? DefaultFillColor => GetDefaultFillColor();

        internal override Color? DefaultBorderColor => GetDefaultBorderColor();

        public eLegendPosition Position => eLegendPosition.Left;

        public EPPlusReadOnlyList<ExcelChartLegendEntry> Entries => null;

        public ExcelTextBody TextBody => _dataTable.TextBody;

        double _dataTableWidth, _columnsWidth;
        internal ChartDataTableRenderer(ChartRenderer svgChart) : base(svgChart)
        {
            _dataTable = svgChart.Chart.PlotArea.DataTable;
            Rectangle = new RectRenderItem(svgChart.Plotarea.Rectangle);
            LeftMargin = RightMargin = 3;   
            TopMargin = BottomMargin = 3;
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
            _maxWidth = svgChart.ChartArea.Rectangle.Width - svgChart.ChartArea.LeftMargin - svgChart.ChartArea.RightMargin;
            _maxHeight = svgChart.ChartArea.Rectangle.Height * (3D / 4D) - svgChart.ChartArea.TopMargin - svgChart.ChartArea.RightMargin; //We set max height to 3/4 of the chart area height. The data table will be placed below the plot area. The remaining 1/4 of the chart area height is reserved for the legend and plot area.
            Rectangle.Width = _maxWidth;   //Set to max. Adjust later to actual width
            Rectangle.Height = _maxHeight; //Set to max. Adjust later to actual height
            var items = new List<List<DrawingTextBody>>();
            var headers = new List<DrawingTextBody>();
            items.Add(headers);

            ExcelDrawingParagraph paragraph;
            if (_dataTable.HasFont)
            {
                paragraph = _dataTable.TextBody.Paragraphs.FirstOrDefault();
            }
            else
            {
                paragraph = null;
            }

            double entryWidth = 0, entryHeight=0;
            var values = svgChart.HorizontalAxis.Axis.GetAxisValues(out _, out _, out _);
            var horizontaValues = GetFormattedValues(svgChart, values);
            foreach (var v in horizontaValues)
            {
                var tb = new DrawingTextBody(RenderContext, svgChart.Chart, Rectangle, true);
                if (paragraph == null)
                {
                    tb.AddParagraph(v);
                }
                else
                {
                    tb.ImportParagraph(paragraph, 0, v);
                }
                headers.Add(tb);
                if (entryWidth < tb.Width)
                {
                    entryWidth = tb.Width;
                }
                if(entryHeight < tb.Height)
                {
                    entryHeight = tb.Height;
                }
                _seriesHeadersMeasure.Add(new TextMeasurement((float)tb.Width, (float)tb.Height));
            }

            DataTableRenderItems.Add(headers);

            _dataTableWidth = GetDataTableWidth(svgChart, _dataTable);

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
                        SeriesIcon.Add(sls);
                    }
                    else if (ct.IsTypeColumn() || ct.IsTypeBar())
                    {
                        LegendIconRenderer.SetBarLegend(ChartRenderer, this, ct, index, pSls, serie, sls, entryWidth, entryHeight, maxIconLength);
                        SeriesIcon.Add(sls);
                    }
                    else if (ct.IsTypePie())
                    {
                        LegendIconRenderer.SetPieLegend(ChartRenderer, this, ct, index, pSls, serie, sls, entryWidth, entryHeight, maxIconLength);
                    }
                    var rows = AddSerieValues(svgChart, _maxWidth, _maxHeight, serie);
                    pSls = sls;
                    var formattedValues = serie.GetValues(false, true);
                    var l=new List<DrawingTextBody>();
                    foreach (var v in formattedValues)
                    {
                        var tb = new DrawingTextBody(RenderContext, ChartRenderer.Chart, Rectangle, true);
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
            _columnsWidth = (_dataTableWidth - (GetLegendColWidth()+LeftMargin+RightMargin)) / headers.Count;
            SetColumnWidth(entryHeight);
        }


        private void SetColumnWidth(double entryWidth)
        {
            var height = GetLegendColHeight(1);
            var r = 1;
            foreach (var lc in SeriesIcon)
            {
                if (_dataTable.ShowKeys)
                {
                    //if (lc.MarkerIcon != null)
                    //{
                    //    lc.MarkerIcon.Height += height;
                    //    if (lc.MarkerBackground != null)
                    //    {
                    //        lc.MarkerBackground.Height += height;
                    //    }
                    //}
                    if(lc.SeriesIcon.Top > lc.Textbox.Top)
                    {
                        var diff = lc.SeriesIcon.Top - lc.Textbox.Top;
                        if(lc.SeriesIcon is LineRenderItem line)
                        {
                            line.Y1 = line.Y2 = height + diff;
                        }
                        else
                        {
                            lc.SeriesIcon.Top = height + diff;
                        }
                        lc.Textbox.Top = height;
                    }
                    else
                    {
                        var diff = lc.Textbox.Top - lc.SeriesIcon.Top;
                        if (lc.SeriesIcon is LineRenderItem line)
                        {
                            line.Y1 = line.Y2 = height;
                        }
                        else
                        {
                            lc.SeriesIcon.Top = height;
                        }
                        lc.Textbox.Top = height + diff;
                    }
                }
                else
                {
                    lc.Textbox.Left = height;
                }
                height += GetLegendColHeight(r++);
            }

            double y = 0;
            r = 0;
            var lcw = GetLegendColWidth();
            foreach (var row in DataTableRenderItems)
            {
                var x = lcw + LeftMargin + RightMargin;
                foreach (var cell in row)
                {
                    cell.Left += x + (_columnsWidth / 2)-(cell.Width / 2);
                    cell.Top += y;
                    x += _columnsWidth;
                }
                y += GetLegendColHeight(r++);
           }
            Rectangle.Height = y;
            Rectangle.Width = lcw + _columnsWidth * DataTableRenderItems[0].Count;
        }

        const float MarginIconText = 1.5f;
        private double GetLegendColWidth()
        {
            if (SeriesIcon.Count == 0) return 0;
            var w = SeriesIcon.Max(x=>x.Textbox.Width);
            if(_dataTable.ShowKeys)
            {
                return w + SeriesIcon.Max(x=>x.SeriesIcon?.Width??0) + MarginIconText;
            }
            return w;
        }
        private double GetLegendColHeight(int row)
        {
            var h = 0D;
            if (row > 0)
            {
                h = SeriesIcon[row-1].Textbox.Height;
                if (_dataTable.ShowKeys)
                {
                    var sih = SeriesIcon[row-1].SeriesIcon.Height;
                    var mih = SeriesIcon[row-1].MarkerIcon?.Height ?? 0D;
                    h = Math.Max(mih, Math.Max(h, sih));
                }
            }
            foreach(var cell in DataTableRenderItems[row])
            {
                if (cell.Height > h) h = cell.Height;
            }
            return h;
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
        private List<DrawingTextBody> AddSerieValues(ChartRenderer svgChart, double maxWidth, double maxHeight, ExcelChartSerie serie)
        {
            var ret = new List<DrawingTextBody>();
            if (string.IsNullOrEmpty(serie.Series))
            {
                var a = new ExcelAddressBase(serie.Series);
                var ws = svgChart.Chart.WorkSheet;
                if (ws != null)
                {
                    var range = ws.Cells[a.Address];
                    foreach (var cell in range)
                    {
                        var tb = new DrawingTextBody(RenderContext, svgChart.Chart, Rectangle, true);
                        tb.AddParagraph(cell.Text);
                        ret.Add(tb);
                    }
                }
            }
            else if (serie.StringLiteralsY != null && serie.StringLiteralsY.Length > 0)
            {
                foreach (var se in serie.StringLiteralsY)
                {
                    var tb = new DrawingTextBody(RenderContext, svgChart.Chart, Rectangle, true);
                    tb.AddParagraph(se);
                    ret.Add(tb);
                }
            }
            else if (serie.NumberLiteralsY != null && serie.NumberLiteralsY.Length > 0)
            {
                foreach (var nl in serie.NumberLiteralsY)
                {
                    var tb = new DrawingTextBody(RenderContext, svgChart.Chart, Rectangle, true);
                    tb.AddParagraph(nl.ToString());
                    ret.Add(tb);
                }
            }

            return ret;
        }

        internal override Color? GetDefaultFillColor()
        {
            return GetDefaultFillColorForElement(ChartElement.DataTable, (int)Chart.Style);
        }

        internal override Color? GetDefaultBorderColor()
        {
            return GetDefaultBorderColorForElement(ChartElement.DataTable, (int)Chart.Style);
        }

        public override void AppendRenderItems(List<Transform> renderItems)
        {
            var group = new GroupRenderItem(ChartRenderer.Bounds);
            group.Name = "DataTable";
            group.Left = ChartRenderer.ChartArea.LeftMargin;
            group.Top = GetTop();
            //Render series header column.
            foreach (var li in SeriesIcon)
            {
                if (_dataTable.ShowKeys)
                {
                    group.ChildObjects.Add(li.SeriesIcon);
                    if (li.MarkerBackground != null) group.ChildObjects.Add(li.MarkerBackground);
                    if (li.MarkerIcon != null) group.ChildObjects.Add(li.MarkerIcon);
                }
                group.ChildObjects.Add(li.Textbox);
            }

            foreach (var row in DataTableRenderItems)
            {
                foreach (var cell in row)
                {
                    group.ChildObjects.Add(cell);
                }
            }
            
            if(_dataTable.ShowOutline)
            {
                CreateOutlineRenderItem(group);
            }
            if(_dataTable.ShowHorizontalBorder)
            {
                CreateHorizontalBorderRenderItem(group);
            }
            if (_dataTable.ShowVerticalBorder)
            {
                CreateVerticalBorderRenderItem(group);
            }
            renderItems.Add(group);
        }

        private void CreateVerticalBorderRenderItem(GroupRenderItem group)
        {
            var path = new PathRenderItem(group);
            var lcw = GetLegendColWidth() + LeftMargin + RightMargin;
            var bottom = DataTableRenderItems[DataTableRenderItems.Count - 1][0].Bottom;
            for (int c=1;c < DataTableRenderItems[0].Count; c++)
            {
                var left = lcw + c*_columnsWidth;
                path.Commands.Add(new PathCommand(PathCommandType.Move, left, 0, left, bottom));
            }
            path.Style.FillOpacity = 100;
            path.Style.FillColor = "none";
            path.Style.SetDrawingPropertiesBorder(ChartRenderer.Theme, _dataTable.Border, ChartRenderer.Chart.StyleManager?.Style?.DataTable.BorderReference.Color, _dataTable.Border.Width > 0, GetDefaultBorderColor, 0.75);
        }

        private PathRenderItem CreateHorizontalBorderRenderItem(GroupRenderItem group)
        {
            var path = new PathRenderItem(group);
            var lcw = GetLegendColWidth() + LeftMargin + RightMargin;
            var right = lcw + _columnsWidth * DataTableRenderItems[0].Count;
            for (var r = 1; r < DataTableRenderItems.Count;r++)
            {
                var top = DataTableRenderItems[r][0].Top;
                path.Commands.Add(new PathCommand(PathCommandType.Move, 0, top, right, top));
            }
            path.Style.FillOpacity = 100;
            path.Style.FillColor = "none";
            path.Style.SetDrawingPropertiesBorder(ChartRenderer.Theme, _dataTable.Border, ChartRenderer.Chart.StyleManager?.Style?.DataTable.BorderReference.Color, _dataTable.Border.Width > 0, GetDefaultBorderColor, 0.75);
            return path;
        }

        private PathRenderItem CreateOutlineRenderItem(GroupRenderItem group)
        {
            var path = new PathRenderItem(group);

            var bl = DataTableRenderItems[DataTableRenderItems.Count - 1][0];
            var fc = DataTableRenderItems[1][0];
            var lcw = GetLegendColWidth() + LeftMargin + RightMargin;
            var right = lcw + _columnsWidth * DataTableRenderItems[0].Count;
            path.Commands.Add(new PathCommand(PathCommandType.Move, lcw, 0, right, 0, right, bl.Bottom, lcw, bl.Bottom));
            path.Commands.Add(new PathCommand(PathCommandType.End));
            path.Commands.Add(new PathCommand(PathCommandType.Move, 0, fc.Top, right, fc.Top, right, bl.Bottom, 0, bl.Bottom));
            path.Commands.Add(new PathCommand(PathCommandType.End));
            path.Style.FillOpacity = 100;
            path.Style.FillColor = "none";            
            path.Style.SetDrawingPropertiesBorder(ChartRenderer.Theme, _dataTable.Border, ChartRenderer.Chart.StyleManager?.Style?.DataTable.BorderReference.Color, _dataTable.Border.Width > 0, GetDefaultBorderColor, 0.75);
            return path;
        }

        private double GetTop()
        {
            var bottom = ChartRenderer.Plotarea.Rectangle.Bottom;
            if (ChartRenderer.HorizontalAxis?.Axis.Deleted == false && (ChartRenderer.HorizontalAxis.Rectangle?.Bottom??0) > bottom)
            {
                bottom = ChartRenderer.HorizontalAxis.Rectangle.Bottom;
            }
            if (ChartRenderer.SecondHorizontalAxis?.Axis.Deleted == false && (ChartRenderer.SecondHorizontalAxis.Rectangle?.Bottom ?? 0) > bottom)
            {
                bottom = ChartRenderer.SecondHorizontalAxis.Rectangle.Bottom;
            }
            return bottom+TopMargin;
        }
    }
}