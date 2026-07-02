using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Graphics;
using EPPlus.Graphics.Geometry;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils.TypeConversion;
using System.Collections.Generic;
using System.Drawing;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class ChartSerieDataLabelRenderer : ChartDrawingObject
    {
        //positioning is handled by parent item via these
        private List<SvgDataLabelPoint> dataLabels = new List<SvgDataLabelPoint>();

        private RenderItem seriesIcon = null;
        private int _serieIndex = -1;
        ExcelDrawingParagraph defaultParagraph;
        BoundingBox plotAreaBounds;
        BoundingBox _defaultMargins;
        ExcelChartSerieDataLabel _dlblSerie;

        internal double rotation = double.NaN;
        internal Graphics.Point rotationPoint = null;

        internal override Color? DefaultFillColor { get; }

        public ChartSerieDataLabelRenderer(ChartRenderer chart, ExcelChartSerieDataLabel dlblSerie, BoundingBox maxBounds, ExcelChartStandardSerie serie, List<object> xValues, List<object> yValues, int index) : base(chart)
        {
            _serieIndex = index;
            _dlblSerie = dlblSerie;
            plotAreaBounds = chart.Plotarea.Group.Bounds;

            DefaultFillColor = Color.Transparent;

            if (dlblSerie.TextBody.Paragraphs.Count != 0)
            {
                defaultParagraph = dlblSerie.TextBody.Paragraphs[0];
            }

            dlblSerie.TextBody.GetInsetsInPoints(out double l, out double top, out double right, out double bottom);
            _defaultMargins = new BoundingBox(l, top, right, bottom);

            if (dlblSerie.DataLabels.Count == 0 && serie.NumberOfItems > 0)
            {
                for (int i = 0; i < serie.NumberOfItems; i++)
                {
                    var yVal = yValues == null ? null : yValues[i];
                    var xVal = xValues == null ? null : xValues[i];
                    AddDatalabel(serie, dlblSerie, xVal, yValues[i], maxBounds);
                }
            }
            else
            {
                int nextIndex = dlblSerie.DataLabels[0].Index;
                var customIndex = 0;
                for (int i = 0; i < serie.NumberOfItems; i++)
                {
                    if (i == nextIndex)
                    {
                        var dataLabel = dlblSerie.DataLabels[customIndex++];
                        var individualIndex = dataLabel.Index;
                        var yVal = yValues == null ? null : yValues[i];
                        var xVal = xValues == null ? null : xValues[i];
                        AddDatalabel(serie, dataLabel, xVal, yVal, maxBounds);

                        if (customIndex < dlblSerie.DataLabels.Count)
                        {
                            nextIndex = dlblSerie.DataLabels[customIndex].Index;
                        }
                    }
                    else
                    {
                        var yVal = yValues == null ? null : yValues[i];
                        var xVal = xValues == null ? null : xValues[i];
                        AddDatalabel(serie, dlblSerie, xVal, yValues[i], maxBounds);
                    }
                }
            }
        }

        private void CreateSeriesIcon(ExcelChartStandardSerie serie, BoundingBox maxBounds)
        {
            if (ChartRenderer.Legend == null)
            {
                seriesIcon = ChartRenderer.GetSeriesIcon(serie, _serieIndex, maxBounds);
            }
            else
            {
                var legendItem = ChartRenderer.Legend;
                var seriesIconOrig = (LineRenderItem)legendItem.SeriesIcon[_serieIndex].SeriesIcon;
                var clonedIcon = seriesIconOrig.Clone();

                clonedIcon.Y1 = 0;
                clonedIcon.Y2 = 0;

                seriesIcon = clonedIcon;
            }
        }

        private RenderItem GetSeriesIcon(ExcelChartStandardSerie serie, BoundingBox maxBounds)
        {
            if(seriesIcon == null)
            {
                CreateSeriesIcon(serie, maxBounds);
            }

            return seriesIcon;
        }

        private void AddDatalabel(ExcelChartStandardSerie serie, ExcelChartDataLabelStandard dataLabel, object xValue, object yValue, BoundingBox maxBounds)
        {
            var newDataLabel = new SvgDataLabelPoint(ChartRenderer, dataLabel);
            newDataLabel.ImportDataLabel(serie, dataLabel, xValue, yValue, defaultParagraph, maxBounds, _defaultMargins);

            if(dataLabel.ShowLegendKey)
            {
                newDataLabel.AddSeriesIcon(GetSeriesIcon(serie, maxBounds));
            }

            dataLabels.Add(newDataLabel);
        }

        internal void SetDimensions(int index, Transform basePoint, Transform endPoint)
        {
            if (dataLabels.Count > index)
            {
                dataLabels[index].SetShapeDimensions(basePoint, endPoint);
            }
        }

        internal void SetParentPoint(BoundingBox parent, int index)
        {
            if (dataLabels.Count > index)
            {
                dataLabels[index].SetParentPoint(parent);
            }
        }

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var plotAreaGroup = new GroupRenderItem(plotAreaBounds);

            plotAreaGroup.Left = plotAreaBounds.Position.X;
            plotAreaGroup.Top = plotAreaBounds.Position.Y;

            if(rotation != double.NaN)
            {
                if(rotationPoint != null)
                {
                    plotAreaGroup.RotationPoint = rotationPoint;
                }
                plotAreaGroup.Rotation = rotation;
            }

            if (_dlblSerie.Fill.IsEmpty == false)
            {
                Rectangle.SetDrawingPropertiesFill(ChartRenderer.Theme, _dlblSerie.Fill, null);
                plotAreaGroup.SetDrawingPropertiesFill(ChartRenderer.Theme, _dlblSerie.Fill, null);
            }

            renderItems.Add(plotAreaGroup);
            for(int i = 0; i< dataLabels.Count; i++) 
            {
                if(rotation != double.NaN)
                {
                    dataLabels[i].CounterRotation = -rotation;
                }
                dataLabels[i].AppendRenderItems(plotAreaGroup.RenderItems);
            }
        }
    }
}
