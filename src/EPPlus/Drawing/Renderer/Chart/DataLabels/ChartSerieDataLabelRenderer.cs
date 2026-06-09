using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Graphics;
using EPPlus.Graphics.Geometry;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Collections.Generic;

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

        public ChartSerieDataLabelRenderer(ChartRenderer chart, ExcelChartSerieDataLabel dlblSerie, BoundingBox maxBounds, ExcelChartStandardSerie serie, List<object> xValues, List<object> yValues, int index) : base(chart)
        {
            _serieIndex = index;
            _dlblSerie = dlblSerie;
            plotAreaBounds = chart.Plotarea.Group.Bounds;

            if (dlblSerie.TextBody.Paragraphs.Count != 0)
            {
                defaultParagraph = dlblSerie.TextBody.Paragraphs[0];
                dlblSerie.TextBody.GetInsetsInPoints(out double l, out double top, out double right, out double bottom);
                _defaultMargins = new BoundingBox(l, top, right, bottom);
            }


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
                //if (xValues != null)
                //{
                    for (int i = 0; i < dlblSerie.DataLabels.Count; i++)
                    {
                        var dataLabel = dlblSerie.DataLabels[i];
                        var yVal = yValues == null ? null : yValues[i];
                        var xVal = xValues == null ? null : xValues[i];
                        AddDatalabel(serie, dataLabel, xVal, yVal, maxBounds);
                    }
                //}
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

        BoundingBox _parentShapeBounds = null;
        Vector2 _startToEndDir = Vector2.Zero;
        /// <summary>
        /// Datapoints can have different shapes. 
        /// Which gives different meaning to positions like'Center' and 'Inside' and 'Outside'
        /// Therefore you have the option to provide the bounds of a shape and its endpoint
        /// </summary>
        /// <param name="parentBounds"></param>
        /// <param name="parentPoint"></param>
        /// <param name="index"></param>
        internal void SetParentShape(BoundingBox parentBounds, BoundingBox shapeEndPoint, int index)
        {
            _parentShapeBounds = parentBounds;
            SetParentPoint(shapeEndPoint, index);
        }

        internal void SetParentVector(BoundingBox parentPoint, int index, Vector2 startToEndDir)
        {
            _startToEndDir = startToEndDir;
            SetParentPoint(parentPoint, index);
        }

        internal void SetParentPoint(BoundingBox parent, int index)
        {
            if (dataLabels.Count > index)
            {
                dataLabels[index].SetParentPoint(parent, _parentShapeBounds, _startToEndDir);
            }
            //dataLabels[index].SetParentPoint(parent);
        }

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var plotAreaGroup = new GroupRenderItem(plotAreaBounds);
            plotAreaGroup.Left = plotAreaBounds.Position.X;
            plotAreaGroup.Top = plotAreaBounds.Position.Y;

            if (_dlblSerie.Fill.IsEmpty == false)
            {
                plotAreaGroup.SetDrawingPropertiesFill(ChartRenderer.Theme, _dlblSerie.Fill, null);
                plotAreaGroup.GroupTransform += $" fill=\"{plotAreaGroup.FillColor}\"";
            }

            renderItems.Add(plotAreaGroup);
            for(int i = 0; i< dataLabels.Count; i++) 
            {
                dataLabels[i].AppendRenderItems(plotAreaGroup.RenderItems);
            }
        }
    }
}
