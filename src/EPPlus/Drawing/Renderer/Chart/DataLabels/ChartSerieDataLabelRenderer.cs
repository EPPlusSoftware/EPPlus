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
        ExcelChartDataLabel _dlbl;

        internal double rotation = double.NaN;
        internal Graphics.Point rotationPoint = null;

        internal override Color? DefaultFillColor { get; }

        double? SummedSeries = null;

        public ChartSerieDataLabelRenderer(ChartRenderer chart, ExcelChartDataLabel dlbl, BoundingBox maxBounds, ExcelChartStandardSerie serie, List<object> xValues, List<object> yValues, int index) : base(chart)
        {
            _serieIndex = index;
            _dlbl = dlbl;
            plotAreaBounds = chart.Plotarea.Group.Bounds;

            DefaultFillColor =  dlbl.Fill != null && dlbl.Fill.Color.IsEmpty == false ? dlbl.Fill.Color : Color.Transparent;


            if(yValues != null && yValues.Count != 0)
            {
                SummedSeries = 0d;
                for (var i = 0; i < yValues.Count; i++)
                {
                    SummedSeries += ConvertUtil.GetValueDouble(yValues[i]);
                }
            }

            if (dlbl.TextBody.Paragraphs.Count != 0)
            {
                defaultParagraph = dlbl.TextBody.Paragraphs[0];
            }

            dlbl.TextBody.GetInsetsInPoints(out double l, out double top, out double right, out double bottom);
            _defaultMargins = new BoundingBox(l, top, right, bottom);

            var dlblSerie = dlbl as ExcelChartSerieDataLabel;
            if (dlblSerie == null || dlblSerie.DataLabels.Count == 0)
            {
                for (int i = 0; i < serie.NumberOfItems; i++)
                {
                    var yVal = yValues == null ? null : yValues[i];
                    var xVal = xValues == null ? null : xValues[i];
                    AddDatalabel(serie, dlbl, xVal, yValues[i], maxBounds);
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
                var seriesIconOrig = legendItem.SeriesIcon[_serieIndex].SeriesIcon;
                var clonedIcon = seriesIconOrig.Clone();

                if(clonedIcon is LineRenderItem lineIcon)
                {
                    lineIcon.Y1 = 0;
                    lineIcon.Y2 = 0;
                }
                else if (clonedIcon is RectRenderItem rectIcon)
                {
                    rectIcon.Left = 0;
                    rectIcon.Top = 0;
                }


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

        private void AddDatalabel(ExcelChartStandardSerie serie, ExcelChartDataLabel dataLabel, object xValue, object yValue, BoundingBox maxBounds)
        {
            var newDataLabel = new SvgDataLabelPoint(ChartRenderer, dataLabel, DefaultFillColor);
            newDataLabel.ImportDataLabel(serie, dataLabel, xValue, yValue, defaultParagraph, maxBounds, _defaultMargins, SummedSeries);

            if(dataLabel.ShowLegendKey)
            {
                newDataLabel.AddSeriesIcon(GetSeriesIcon(serie, maxBounds));
            }

            dataLabels.Add(newDataLabel);
        }

        internal void SetDimensions(int index, Transform basePoint, Transform endPoint, BoundingBox maxBoundsPieSlice = null)
        {
            if (dataLabels.Count > index)
            {
                dataLabels[index].SetShapeDimensions(basePoint, endPoint, maxBoundsPieSlice);
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

            if (_dlbl.Fill.IsEmpty == false)
            {
                Rectangle.SetDrawingPropertiesFill(ChartRenderer.Theme, _dlbl.Fill, null);
                plotAreaGroup.SetDrawingPropertiesFill(ChartRenderer.Theme, _dlbl.Fill, null);
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
