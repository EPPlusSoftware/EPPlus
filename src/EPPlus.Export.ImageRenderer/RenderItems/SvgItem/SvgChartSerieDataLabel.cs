using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Collections.Generic;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgChartSerieDataLabel : DrawingObjectNoBounds
    {
        //positioning is handled by parent item via these
        internal List<SvgGroupItem> groupItems = new List<SvgGroupItem>();
        private List<SvgDataLabelPoint> dataLabels = new List<SvgDataLabelPoint>();

        private RenderItem seriesIcon = null;
        private int _serieIndex = -1;
        ExcelDrawingParagraph defaultParagraph;
        BoundingBox plotAreaBounds;
        BoundingBox _defaultMargins;
        ExcelChartSerieDataLabel _dlblSerie;

        public SvgChartSerieDataLabel(SvgChart chart, ExcelChartSerieDataLabel dlblSerie, BoundingBox maxBounds, ExcelChartStandardSerie serie, List<object> xValues, List<object> yValues, int index) : base(chart)
        {
            _serieIndex = index;
            _dlblSerie = dlblSerie;
            plotAreaBounds = chart.Plotarea.Rectangle.Bounds;

            if (dlblSerie.TextBody.Paragraphs.Count != 0)
            {
                defaultParagraph = dlblSerie.TextBody.Paragraphs[0];
                dlblSerie.TextBody.GetInsetsInPoints(out double l, out double top, out double right, out double bottom);
                _defaultMargins = new BoundingBox(l, top, right, bottom);
            }


            if (dlblSerie.DataLabels.Count == 0 && serie.NumberOfItems > 0)
            {
                //if (xValues != null)
                //{
                    for (int i = 0; i < serie.NumberOfItems; i++)
                    {
                        var yVal = yValues == null ? null : yValues[i];
                        var xVal = xValues == null ? null : xValues[i];
                        AddDatalabel(chart, serie, dlblSerie, xVal, yValues[i], maxBounds);
                    }
                //}
            }
            else
            {
                //if (xValues != null)
                //{
                    for (int i = 0; i < dlblSerie.DataLabels.Count; i++)
                    {
                        var dataLabel = dlblSerie.DataLabels[i];
                        var yVal = yValues == null ? null : yValues[i];
                        var xVal = xValues == null ? null : xValues[i];
                        AddDatalabel(chart, serie, dataLabel, xVal, yVal, maxBounds);
                    }
                //}
            }
        }

        private void CreateSeriesIcon(SvgChart chart, ExcelChartStandardSerie serie, BoundingBox maxBounds)
        {
            if (chart.Legend == null)
            {
                seriesIcon = chart.GetSeriesIcon(serie, _serieIndex, maxBounds);
            }
            else
            {
                var legendItem = chart.Legend;
                var seriesIconOrig = (SvgRenderLineItem)legendItem.SeriesIcon[_serieIndex].SeriesIcon;
                var clonedIcon = seriesIconOrig.Clone(chart);

                clonedIcon.Y1 = 0;
                clonedIcon.Y2 = 0;

                seriesIcon = clonedIcon;
            }
        }

        private RenderItem GetSeriesIcon(SvgChart chart, ExcelChartStandardSerie serie, BoundingBox maxBounds)
        {
            if(seriesIcon == null)
            {
                CreateSeriesIcon(chart, serie, maxBounds);
            }

            return seriesIcon;
        }

        private void AddDatalabel(SvgChart chart, ExcelChartStandardSerie serie, ExcelChartDataLabelStandard dataLabel, object xValue, object yValue, BoundingBox maxBounds)
        {
            var newDataLabel = new SvgDataLabelPoint(chart, dataLabel);
            newDataLabel.ImportDataLabel(chart, serie, dataLabel, xValue, yValue, defaultParagraph, maxBounds, _defaultMargins);

            if(dataLabel.ShowLegendKey)
            {
                newDataLabel.AddSeriesIcon(GetSeriesIcon(chart, serie, maxBounds));
            }

            dataLabels.Add(newDataLabel);
        }

        BoundingBox _parentShapeBounds = null;
        Graphics.Math.Vector2 _startToEndDir = Graphics.Math.Vector2.Zero;

        internal void SetParentVector(BoundingBox parentPoint, int index, Graphics.Math.Vector2 startToEndDir)
        {
            _startToEndDir = startToEndDir;
            SetParentPoint(parentPoint, index);
        }

        internal void SetParentPoint(BoundingBox parent, int index)
        {
            if (dataLabels.Count > index)
            {
                dataLabels[index].SetParentPoint(parent, _startToEndDir);
            }
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var plotAreaGroup = new SvgGroupItem(DrawingRenderer, plotAreaBounds);

            if(_dlblSerie.Fill.IsEmpty == false)
            {
                plotAreaGroup.SetDrawingPropertiesFill(_dlblSerie.Fill, null);

                plotAreaGroup.GroupTransform += $" fill=\"{plotAreaGroup.FillColor}\"";
            }

            renderItems.Add(plotAreaGroup);
            for(int i = 0; i< dataLabels.Count; i++) 
            {
                dataLabels[i].AppendRenderItems(renderItems);
            }
            renderItems.Add(new SvgEndGroupItem(DrawingRenderer, plotAreaBounds));
        }
    }
}
