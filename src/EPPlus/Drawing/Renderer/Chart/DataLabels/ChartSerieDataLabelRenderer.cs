using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Graphics;
using EPPlus.Graphics.Geometry;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Renderer.Chart;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Net;
using System.Security;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class ChartSerieDataLabelRenderer : ChartDrawingObjectWithBackground
    {
        //positioning is handled by parent item via these
        private List<SvgDataLabelPoint> dataLabels = new List<SvgDataLabelPoint>();

        private RenderItem seriesIcon = null;
        private int _serieIndex = -1;
        private int _origIndex = -1;
        ExcelDrawingParagraph defaultParagraph;
        BoundingBox plotAreaBounds;
        BoundingBox _defaultMargins;
        ExcelChartDataLabel _dlbl;

        internal double rotation = double.NaN;
        internal Graphics.TranformPoint rotationPoint = null;

        internal override Color? DefaultFillColor { get; }

        internal override Color? DefaultBorderColor => null;

        double? SummedSeries = null;

        GroupRenderItem DatalabelsGroup;

        public ChartSerieDataLabelRenderer(ChartRenderer chart, ExcelChartDataLabel dlbl, BoundingBox maxBounds, ExcelChartStandardSerie serie, List<object> xValues, List<object> yValues, int index) : base(chart)
        {
            DatalabelsGroup = new GroupRenderItem(ChartRenderer.Bounds);
            DatalabelsGroup.Name = "Dlbl_SeriesGroup";

            _serieIndex = index;
            _origIndex = index;
            _dlbl = dlbl;
            plotAreaBounds = chart.Plotarea.Group;

            DefaultFillColor =  dlbl.Fill != null && dlbl.Fill.Color.IsEmpty == false ? dlbl.Fill.Color : Color.Transparent;

            Rectangle.Name = "SerieDataLabel";


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

            dlbl.TextBody.GetInsetsInPointsNullable(out double? nullL, out double? nullT, out double? nullR, out double? nullB);

            //Datalabels have different default margins than standard Rectangle shape
            double l = nullL ?? 3.1181102362d;
            double t = nullT ?? 1.4173228346d;
            double r = nullR ?? 3.1181102362d;
            double b = nullB ?? 1.4173228346d;

            _defaultMargins = new BoundingBox(l, t, r, b);

            var dlblSerie = dlbl as ExcelChartSerieDataLabel;
            if (dlblSerie == null || dlblSerie.DataLabels.Count == 0)
            {
                for (int i = 0; i < serie.NumberOfItems; i++)
                {
                    var yVal = yValues == null ? null : yValues[i];
                    var xVal = xValues == null ? null : xValues[i];
                    if (dlblSerie == null)
                    {
                        AddDatalabel(serie, dlbl, xVal, yValues[i], maxBounds);
                    }
                    else
                    {
                        AddDatalabel(serie, dlblSerie, xVal, yValues[i], maxBounds);
                    }
                    //Bit strange but in e.g. pie charts each datapoint counts as a new series for the purposes of legendIcons etc.
                    _serieIndex++;
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
        
        private void CreateSeriesIcon(ExcelChartStandardSerie serie, BoundingBox maxBounds, double entryHeight)
        {
            try
            {
                if (ChartRenderer.Legend == null)
                {
                    seriesIcon = LegendIconRenderer.GetSeriesIcon(ChartRenderer, maxBounds, serie, _serieIndex, entryHeight);
                }
                else
                {
                    var legendItem = ChartRenderer.Legend;
                    RenderItem seriesIconOrig;
                    if (_serieIndex >= legendItem.SeriesIcon.Count)
                    {
                        seriesIconOrig = legendItem.SeriesIcon[legendItem.SeriesIcon.Count-1].SeriesIcon;
                    }
                    else
                    {
                        seriesIconOrig = legendItem.SeriesIcon[_serieIndex].SeriesIcon;
                    }

                    var clonedIcon = seriesIconOrig.Clone();

                    if (seriesIconOrig.Style.FillColor == null && seriesIconOrig.Style.GradientFill != null)
                    {
                        clonedIcon.Style.GradientFill = seriesIconOrig.Style.GradientFill;
                    }

                    if (clonedIcon is LineRenderItem lineIcon)
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
            catch(Exception e)
            {
                throw new Exception("Something went wrong in datalabel renderer SeriesIcon", e);
            }
        }

        private RenderItem GetSeriesIcon(ExcelChartStandardSerie serie, BoundingBox maxBounds, double entryHeight)
        {
            //We MUST create a new icon per series. For pie chart each data point is a new series
            //Therefore check if _origIndex matches
            if (seriesIcon == null || _origIndex != _serieIndex)
            {
                CreateSeriesIcon(serie, maxBounds, entryHeight);
                return seriesIcon;
            }
            else
            {
                return seriesIcon.Clone();
            }
        }

        private void AddDatalabel(ExcelChartStandardSerie serie, ExcelChartDataLabel dataLabel, object xValue, object yValue, BoundingBox maxBounds)
        {
            //We don't add it if it's specified to be deleted
            if (dataLabel.HasDeleteNode == false)
            {
                var newDataLabel = new SvgDataLabelPoint(ChartRenderer, dataLabel, DefaultFillColor);

                var dlblName = $"DatalabelPoint_{dataLabels.Count}";
                newDataLabel.ImportDataLabel(serie, dataLabel, xValue, yValue, defaultParagraph, maxBounds, _defaultMargins, SummedSeries);
                newDataLabel.DataLabelPointContainer.Name = dlblName;

                if (dataLabel.ShowLegendKey)
                {
                    newDataLabel.AddSeriesIcon(GetSeriesIcon(serie, maxBounds, newDataLabel.Rectangle.Height)); //TODO: Check if Rectangle.Height matches entry height
                }

                dataLabels.Add(newDataLabel);
            }
        }

        internal double GetTextWidth(int index)
        {
            if (dataLabels.Count > index)
            {
                return dataLabels[index].GlobalTextBounds.Width;
            }
            throw new InvalidOperationException("Index larger than collection");
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

        public override void AppendRenderItems(List<Transform> renderItems)
        {
            DatalabelsGroup.Left = ChartRenderer.Plotarea.Group.Left;
            DatalabelsGroup.Top = ChartRenderer.Plotarea.Group.Top;

            if (double.IsNaN(rotation) == false)
            {
                if (rotationPoint != null)
                {
                    DatalabelsGroup.RotationPoint = rotationPoint;
                }
                DatalabelsGroup.Rotation = rotation;
            }

            if (_dlbl.Fill.IsEmpty == false)
            {
                Rectangle.Style.SetDrawingPropertiesFill(ChartRenderer.Theme, _dlbl.Fill, null);
                DatalabelsGroup.Style.SetDrawingPropertiesFill(ChartRenderer.Theme, _dlbl.Fill, null);
            }

            for (int i = 0; i < dataLabels.Count; i++)
            {
                if (double.IsNaN(rotation) == false)
                {
                    dataLabels[i].CounterRotation = -rotation;
                }

                dataLabels[i].PrepareForRenderAndAddToParent(DatalabelsGroup);
            }

            //With semi-final positions check if any of the BestFit ones overlap
            for (int i = 0; i < dataLabels.Count; i++)
            {
                var currentLabel = dataLabels[i];
                //var currentLabelRect = new Rect(c)

                if(currentLabel.labelPosition == eLabelPosition.BestFit)
                {
                    //var limitBot = currentLabel.limitBottom;
                    //var limitTop = currentLabel.limitTop;
                    ////Adjust txtbox to not go outside inner maxBounds
                    //if (currentLabel.GlobalTextBounds.Bottom > limitBot)
                    //{
                    //    var overlap = limitBot - currentLabel.GlobalTextBounds.Bottom;
                    //    currentLabel.DataLabelPointContainer.Top -= overlap + currentLabel.BottomMargin;
                    //}

                    //for (int j = 0; j < i; j++)
                    //{
                    //    var bbCurrent = currentLabel.GlobalTextBounds;
                    //    var bbCompare = dataLabels[j].GlobalTextBounds;

                    //    if (Transform.Intersects(bbCurrent, bbCompare))
                    //    {
                    //        var adjustVector = Transform.GetIntersectAdjustVector(bbCurrent, bbCompare);

                    //        if((adjustVector.X == 0 && adjustVector.Y == 0) == false)
                    //        {
                    //            currentLabel.DataLabelPointContainer.Left += adjustVector.X;
                    //            currentLabel.DataLabelPointContainer.Top += adjustVector.Y;
                    //        }
                    //    }
                    //}
                }
            }

            renderItems.Add(DatalabelsGroup);
        }

        internal override Color? GetDefaultFillColor()
        {
            return DefaultFillColor;
        }

        internal override Color? GetDefaultBorderColor()
        {
            return DefaultBorderColor;
        }
    }
}
