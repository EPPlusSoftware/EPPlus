using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Utils;
using EPPlus.Graphics;
using EPPlus.Graphics.Geometry;
using EPPlus.Graphics.TransformPrimitives;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Renderer.Chart.ChartTypeDrawers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils.Drawing;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;

namespace EPPlus.Export.ImageRenderer.Svg.Chart.ChartTypeDrawers
{
    internal class PieChartTypeDrawer : ChartTypeDrawer
    {
        List<ChartSerieDataLabelRenderer> serieDataLabels = new List<ChartSerieDataLabelRenderer>();
        List<object> catValues = new List<object>();
        List<object> valValues = new List<object>();
        List<double> _serieValuesAsDoubles = new List<double>();
        List<List<BoundingBox>> dataPointsPerSerie = new List<List<BoundingBox>>();
        double _totalOfSerieValues = 0;

        List<PieSliceRenderItem> Slices = new List<PieSliceRenderItem>();

        GroupRenderItem _groupItem;

        double _startDegrees = -90d;
        int _pieExplosionPercent = 0;

        double _sliceScaleFactor = 1.0;
        int _serCounter = 0;

        /// <summary>
        /// Radius in points
        /// </summary>
        double _radius;

        TranformPoint _circleCenter;

        internal override Color? DefaultFillColor => null;

        internal override Color? DefaultBorderColor => null;

        PieTransform pieTransform;

        public PieChartTypeDrawer(ChartRenderer chart, ExcelPieChart chartType) : base(chart, chartType)
        {
            //Moved to draw series
        }

        void RenderDebugEllipse()
        {
            var circ = new EllipseRenderItem(ChartRenderer.Plotarea.Rectangle);

            circ.Left = _circleCenter.Left;
            circ.Top = _circleCenter.Top;

            circ.Rx = _radius;
            circ.Ry = _radius;

            circ.Cx = _circleCenter.Left;
            circ.Cy = _circleCenter.Top;

            circ.Style.FillColor = "transparent";
            circ.Style.FillOpacity = 0.3d;
            circ.Style.BorderColor = "purple";
            circ.Style.BorderWidth = 10;

            _groupItem.ChildObjects.Add(circ);
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

            List<double> valPercents = new List<double>();
            

            for (int i = 0; i < _serieValuesAsDoubles.Count; i++)
            {   //Calculate how many percent of the pie this slice is
                var valPercent = _serieValuesAsDoubles[i] / _totalOfSerieValues;
                valPercents.Add(valPercent);
            }


            pieTransform = new PieTransform(_radius);
            pieTransform.GenerateSlicesAndClearPrevious(valPercents);

            pieTransform.Parent = _circleCenter.Parent;
            pieTransform.Position = _circleCenter.Position;
            pieTransform.Name = "Pie_Transform";

            for (int i = 0; i < valPercents.Count; i++)
            {
                PieSliceRenderItem renderSlice = new PieSliceRenderItem(ChartRenderer, _groupItem, pieTransform.TransformSlices[i], pieTransform.Slices.Percentages[i], pieTransform, _startDegrees, prevDegrees);
                prevDegrees = renderSlice.Degrees + prevDegrees;
                Slices.Add(renderSlice);
            }
        }

        double _shrinkFactor = 1;

        void CalculateLocalCenterAndRadius()
        {
            if (_shouldShrink == true)
            {
                _shrinkFactor = 0.8632625d;
            }

            _circleCenter = new TranformPoint();
            _circleCenter.Parent = _groupItem.Parent;
            _circleCenter.Left = (ChartRenderer.Plotarea.Rectangle.Width / 2d);
            _circleCenter.Top = (ChartRenderer.Plotarea.Rectangle.Height / 2d);

            _groupItem.RotationPoint = _circleCenter;


            _radius = Math.Min(_circleCenter.Left * _shrinkFactor, _circleCenter.Top * _shrinkFactor);
        }

        bool _shouldShrink = false;

        List<ExcelChartDataLabel> _serieLabels;

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
                    if (chartType.HasDataLabel || serie.HasDataLabel)
                    {
                        ChartSerieDataLabelRenderer datalabel;
                        if (serie.HasDataLabel)
                        {
                            //_serieLabels.Add(serie.DataLabel);
                            datalabel = new ChartSerieDataLabelRenderer(ChartRenderer, serie.DataLabel, ChartRenderer.Bounds, serie, catValues, valValues, _serCounter);
                            if ((serie.DataLabel.Position == eLabelPosition.BestFit || serie.DataLabel.Position == eLabelPosition.OutEnd) && _shouldShrink == false)
                            {
                                _shouldShrink = true;
                            }
                        }
                        else
                        {
                            //_serieLabels.Add(chartType.DataLabel);
                            datalabel = new ChartSerieDataLabelRenderer(ChartRenderer, chartType.DataLabel, ChartRenderer.Bounds, serie, catValues, valValues, _serCounter);
                            if ((chartType.DataLabel.Position == eLabelPosition.BestFit || chartType.DataLabel.Position == eLabelPosition.OutEnd) && _shouldShrink == false)
                            {
                                _shouldShrink = true;
                            }
                        }

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
                _totalOfSerieValues += dValue;
            }
        }

        private void UpdateSlice(ExcelPieChart chartType, ExcelPieChartSerie serie, int seriesCount, int position)
        {
            int explosion = 0;
            if (serie.DataPoints.ContainsKey(position))
            {
                explosion = serie.DataPoints[position].Explosion;
            }

            Slices[position].ImportPathData(
                ChartRenderer.Plotarea.Rectangle, ChartRenderer.Bounds,
                _sliceScaleFactor, explosion, _pieExplosionPercent, position);

            Slices[position].ImportStlyeInfo(serie, chartType, position);
            Slices[position].AppendGroupItem(_groupItem);
        }

        private ExcelChartSeries<ExcelPieChartSerie> _chartSeries;

        internal override void DrawSeries()
        {
            _groupItem = new GroupRenderItem(ChartRenderer.Plotarea.Group);

            Rectangle.Name = "ChartDrawer";
            _groupItem.Name = "OuterGroupChartDrawer";

            var chartType = (ExcelPieChart)_chartType;

            //Read and set Starting angle offset as a rotation on the container
            //This way no rotation messes with the other calculations
            var angleOffset = double.IsNaN(chartType.FirstSliceAngle) ? 0 : chartType.FirstSliceAngle;
            _groupItem.GroupRotation = angleOffset;

             _chartSeries = chartType.Series;

            LoadSeriesValues(chartType);

            CalculateLocalCenterAndRadius();
            if (serieDataLabels.Count > 0)
            {
                serieDataLabels[0].rotation = angleOffset;
                serieDataLabels[0].rotationPoint = _groupItem.RotationPoint;
            }

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

            int dataPointCount = 0;
            if (catValues != null)
            {
                if (valValues != null)
                {
                    dataPointCount = Math.Max(catValues.Count, valValues.Count);
                }
                else
                {
                    dataPointCount = catValues.Count;
                }
            }
            else
            {
                if (valValues != null)
                {
                    dataPointCount = valValues.Count;
                }
            }

            for (int i = 0; i < chartType.Series.Count; i++)
            {
                var serie = (ExcelPieChartSerie)chartType.Series[i];

                List<BoundingBox> DataLabelGlobalOriginPoints = new();
                List<Vector2> VectorsCenterToMidPointPerSlice = new();

                //Excel ignores series beyond the first for pie chart visualization
                if (i == 0)
                {
                    for (var j = 0; j < Slices.Count; j++)
                    {
                        //Update the initialized slice with path, style and group data
                        UpdateSlice(chartType, serie, dataPointCount, j);

                        if (chartType.HasDataLabel || serie.HasDataLabel)
                        {

                            var innerGroup = Slices[j].GetInnerGroupWithTransformOriginTranslated();
                            //Get the global position of the inner items (innerGroup the parent of itemGroup has already had its position set correctly)
                            var dlblBounds = new BoundingBox(innerGroup.LocalPosition.X, innerGroup.LocalPosition.Y, Rectangle.Width, Rectangle.Height);

                            var ctrToMid = Slices[j].GetWholeVectorCenterToMid();
                            var startPt = new Transform();
                            startPt.Parent = innerGroup.Parent;

                            //This technically places the debug points in wrong position when shrinking but the result looks more similar to excel
                            //innerGroup.Position = innerGroup.Position - ((ctrToMid * -1) * (_shrinkFactor - 1)*_sliceScaleFactor);

                            startPt.Position = innerGroup.Position + (ctrToMid * -1);


                            var bounds = Slices[j].GetBounds();

                            serieDataLabels[i].SetDimensions(j, startPt, innerGroup, bounds);
                        }
                    }
                }
            }

            //RenderDebugEllipse();
            ChartAreaRenderItems.Add(_groupItem);
        }

        public override void AppendRenderItems(List<Transform> renderItems)
        {
            if (serieDataLabels != null && serieDataLabels.Count > 0)
            {
                serieDataLabels[0].AppendRenderItems(ChartRenderer.RenderItems);
            }
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
