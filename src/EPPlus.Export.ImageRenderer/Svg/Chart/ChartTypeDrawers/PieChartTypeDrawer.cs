using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace EPPlus.Export.ImageRenderer.Svg.Chart.ChartTypeDrawers
{
    internal class PieChartTypeDrawer : ChartTypeDrawer
    {
        List<SvgChartSerieDataLabel> serieDataLabels = new List<SvgChartSerieDataLabel>();
        List<object> catValues = new List<object>();
        List<object> valValues = new List<object>();
        List<double> _serieValuesAsDoubles = new List<double>();
        List<List<BoundingBox>> dataPointsPerSerie = new List<List<BoundingBox>>();
        double _totalOfSerieValues = 0;

        List<SvgPieSlice> Slices = new List<SvgPieSlice>();

        SvgGroupItemNew _groupItem;

        double _startDegrees = -90d;
        int _pieExplosionPercent = 0;

        double _sliceScaleFactor = 1.0;
        int _serCounter = 0;

        /// <summary>
        /// Radius in points
        /// </summary>
        double _radius;

        Point _circleCenter;

        public PieChartTypeDrawer(SvgChart chart, ExcelPieChart chartType) : base(chart, chartType)
        {
            //Moved to draw series
        }

        void RenderDebugEllipse()
        {
            var circ = new SvgRenderEllipseItem(ChartRenderer, _svgChart.Plotarea.Rectangle.Bounds);

            circ.Bounds.Left = _circleCenter.Left;
            circ.Bounds.Top = _circleCenter.Top;

            circ.Rx = _radius;
            circ.Ry = _radius;

            circ.Cx = _circleCenter.Left;
            circ.Cy = _circleCenter.Top;

            circ.FillColor = "transparent";
            circ.FillOpacity = 0.3d;
            circ.BorderColor = "purple";
            circ.BorderWidth = 10;

            _groupItem.AddChildItem(circ);
        }

        Coordinate CalculateLocalPointOnCircle(double degrees)
        {
            var angleRadians = MConverter.DegreesToRadians(degrees);

            var xPoint = _circleCenter.Left + (_radius * Math.Cos(angleRadians));
            var yPoint = _circleCenter.Top + (_radius * Math.Sin(angleRadians));

            return new Coordinate(xPoint, yPoint);
        }

        void InitializeSlices()
        {
            //The angle of the previous slice
            //(or the 90 degree offset in the first slice)
            var prevDegrees = _startDegrees;

            for (int i = 0; i < valValues.Count; i++)
            {   //Calculate how many percent of the pie this slice is
                var valPercent = _serieValuesAsDoubles[i] / _totalOfSerieValues;
                //Create and add slice
                SvgPieSlice slice = new SvgPieSlice(ChartRenderer, _groupItem.Bounds, _circleCenter, _radius, valPercent, prevDegrees);
                Slices.Add(slice);

                //Next slice will need to be calculated starting from the degrees of this slice
                prevDegrees = slice.Degrees + prevDegrees;
            }
        }

        void CalculateLocalCenterAndRadius()
        {
            _circleCenter = new Point();
            _circleCenter.Parent = _groupItem.TranslationOffset;
            _circleCenter.Left = _svgChart.Plotarea.Rectangle.Bounds.Width / 2;
            _circleCenter.Top = _svgChart.Plotarea.Rectangle.Bounds.Height / 2;

            _groupItem.RotationPoint = _circleCenter;

            _radius = Math.Min(_circleCenter.Left, _circleCenter.Top);
        }

        void LoadSeriesValues(ExcelPieChart chartType)
        {
            //Load series values
            foreach (ExcelPieChartSerie serie in chartType.Series)
            {
                //Excel allows further series on a pie chart but ignores them for visualization
                if(_serCounter == 0)
                {

                    valValues = LoadSeriesValues(serie.Series, serie.NumberLiteralsY, serie.StringLiteralsY);
                    catValues = LoadSeriesValues(serie.XSeries, serie.NumberLiteralsX, serie.StringLiteralsX);

                    //Pie explosion
                    _pieExplosionPercent = serie.Explosion == int.MinValue ? 0 : serie.Explosion;

                    //Add Datalabel
                    if (serie.HasDataLabel)
                    {
                        var datalabel = new SvgChartSerieDataLabel(_svgChart, serie.DataLabel, _svgChart.Bounds, serie, catValues, valValues, _serCounter);
                        serieDataLabels.Add(datalabel);
                    }
                }

                _serCounter++;
            }

            ConvertSerieValuesToDoubles();
        }

        void ConvertSerieValuesToDoubles()
        {
            for (int i = 0; i < valValues.Count; i++)
            {
                var origValue = valValues[i];
                var dValue = ConvertUtil.GetValueDouble(origValue, false, true);
                if (double.IsNaN(dValue))
                {
                    //Ignore values that are NAN. Possibly we should throw here but Excel simply seems to skip it.
                    continue;
                }
                _serieValuesAsDoubles.Add(dValue);
                _totalOfSerieValues += _serieValuesAsDoubles[i];
            }
        }

        private void UpdateSlice(ExcelPieChart chartType, ExcelPieChartSerie serie, int seriesCount, int position)
        {
            var dataPoint = serie.DataPoints[position];

            Slices[position].ImportPathData(
                _svgChart.Plotarea.Rectangle.Bounds, _svgChart.Bounds, 
                _sliceScaleFactor, dataPoint.Explosion, _pieExplosionPercent, position);

            Slices[position].ImportStlyeInfo(dataPoint, chartType);
            Slices[position].AppendGroupItem(_groupItem);
        }

        internal override void DrawSeries()
        {
            _groupItem = new SvgGroupItemNew(ChartRenderer, 0, 0);
            _groupItem.TranslationOffset.Left = _svgChart.Plotarea.Rectangle.Left;
            _groupItem.TranslationOffset.Top = _svgChart.Plotarea.Rectangle.Top;

            Bounds.Name = "ChartDrawer";

            _groupItem.Bounds.Name = "OuterGroupChartDrawer";

            //_groupitem.bounds.parent = _groupitem.translationoffset;

            var chartType = (ExcelPieChart)_chartType;

            //Read and set Starting angle offset as a rotation on the container
            //This way no rotation messes with the other calculations
            var angleOffset = double.IsNaN(chartType.FirstSliceAngle) ? 0 : chartType.FirstSliceAngle;
            _groupItem.Rotation = angleOffset;

            LoadSeriesValues(chartType);
            CalculateLocalCenterAndRadius();
            InitializeSlices();

            //How much to scale each slice due to pie explosion
            //Essentially we start at 100% (100/100)
            //And then scale down by adding the pie explosionPercent (0-400)
            // 100/200 -> 0.5, 100/400 -> 0.2 etc. 
            _sliceScaleFactor = 100d / (_pieExplosionPercent + 100d);

            if (_sliceScaleFactor != 1)
            {
                //Small adjustment. Unsure why but closer results
                //Could be Excel pixel rounding or 2px border buffer
                _sliceScaleFactor += 0.02d;
            }

            int count = 0;
            if (catValues != null)
            {
                if (valValues != null)
                {
                    count = Math.Min(catValues.Count, valValues.Count);
                }
                else
                {
                    count = catValues.Count;
                }
            }
            else
            {
                if (valValues != null)
                {
                    count = valValues.Count;
                }
            }

            for (int i = 0; i < chartType.Series.Count; i++)
            {
                var serie = (ExcelPieChartSerie)chartType.Series[i];

                List<BoundingBox> DataLabelGlobalOriginPoints = new();
                List<EPPlus.Graphics.Math.Vector2> VectorsCenterToMidPointPerSlice = new();

                //Excel ignores series beyond the first for pie chart visualization
                if (i == 0)
                {
                    for (var j = 0; j < count; j++)
                    {
                        //Update the initialized slice with path, style and group data
                        UpdateSlice(chartType, serie, count, j);

                        if(serie.HasDataLabel)
                        {
                            //Get the inner item local coordinates
                            var itemGroup = Slices[j].GetInnerItemGroup();
                            var outer = Slices[j].GetOuterMidpointInGlobalCoords();
                            ////Get the relevant point local coordinates (to the above)
                            //var slicePos = Slices[j].GetInnerGroupTransformOrigin();

                            //Get the global position of the inner items (innerGroup the parent of itemGroup has already had its position set correctly)
                            var dlblBounds = new BoundingBox(itemGroup.TransformOrigin.X, itemGroup.TransformOrigin.Y, Bounds.Width, Bounds.Height);
                            //var dlblBounds2 = new BoundingBox(outer.X, outer.Y);
                            ////Add the origin point position
                            //dlblBounds.Left += slicePos.X;
                            //dlblBounds.Top += slicePos.Y;

                            ////we now have the start point of the slice in global coords without the translation
                            ////Add the same translation
                            //dlblBounds.Translate()

                            //var vector = Slices[j].GetWholeVectorCenterToMid();
                            serieDataLabels[i].SetParentVector(dlblBounds, j, Slices[j].GetWholeVectorCenterToMid());
                        }
                        ////Add datalabel to slice
                        //if (serie.HasDataLabel)
                        //{
                            
                        //    //DataLabelGlobalOriginPoints.Add();
                        //    //serieDataLabels[i].AppendRenderItems(Slices[j].);
                        //    //var slicePos = Slices[j].GetInnerGroupTransformOrigin();
                        //    //var bounds = Slices[j].GetInnerGroupBounds();
                        //    //BoundingBox dp = new BoundingBox(slicePos.X, slicePos.Y, 0, 0);

                        //serieDataLabels[i].SetParentPoint(Slices[j].GetInnerGroupBounds(), j);
                        //    //serieDataLabels[i].SetParentVector(Slices[j].GetInnerGroupBounds(), j, Slices[j].GetWholeVectorCenterToMid());
                        //    //serieDataLabels[i].SetParentShape(Slices[j].ExtremePoints, dp, j);
                        //}
                    }
                }
            }

            //RenderDebugEllipse();

            RenderItems.Add(_groupItem);
            //Date series labels
            foreach (var dataLabel in serieDataLabels)
            {
                dataLabel.AppendRenderItems(RenderItems);
            }
        }
    }
}
