/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/10/2022         EPPlus Software AB           EPPlus v6
 *************************************************************************************************/
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.DigitalSignatures;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal class SvgTrendline : SvgChartObject
    {
        private ExcelChartTrendline _trendline;
        private double[] _ySerie;
        private List<object> _xSerie;
        private SvgChart _svgChart;
        private ExcelChart _chartType;
        private bool _useSecondaryAxis;
        private int _serieCount, _seriePos;
        public SvgTrendline(SvgChart svgChart, ExcelChartTrendline trendline, List<object> xSerie, List<object> ySerie, ExcelChart chartType, int seriePos) : base(svgChart)
        {
            _svgChart = svgChart;
            _chartType = chartType;
            _trendline = trendline;
            _xSerie = xSerie;
            _ySerie = ySerie.Select(y => ConvertUtil.GetValueDouble(y)).ToArray();
            _useSecondaryAxis = chartType.UseSecondaryAxis;
            _serieCount = _chartType.Series.Count;
            _seriePos = seriePos;
            double m, b;
            switch (trendline.Type) 
            {
                case eTrendLine.Linear:
                    CalculateLinear();
                    Coordinates.Add(new Coordinate(0, GetLinearValueAtPosition(1)));
                    Coordinates.Add(new Coordinate(_xSerie.Count-1, GetLinearValueAtPosition(_xSerie.Count)));
                    break;
                case eTrendLine.Exponential:
                    CalculateExponential();
                    m = Coefficients[0];
                    b = Coefficients[1];
                    CreateCoordinates(x => b * Math.Exp(m * x));
                    break;
                case eTrendLine.Logarithmic:
                    CalculateLogarithmic();
                    m = Coefficients[0];
                    b = Coefficients[1];
                    CreateCoordinates(x => m * Math.Log(x) + b);
                    break;
                case eTrendLine.Polynomial:
                    CalculatePolynomial();
                    CreateCoordinates(x=> PredictPolynomial(x));
                    break;
                case eTrendLine.Power:
                    CalculatePower();

                    m = Coefficients[0];
                    b = Coefficients[1];
                    CreateCoordinates(x => b * Math.Pow(x, m));
                    break;
                case eTrendLine.MovingAverage:
                    CalculateMoveingAverage();

                    for (int i = ((int)_trendline.Period) - 1; i < _xSerie.Count; i++)
                    {
                        var x = i;
                        var y = GetMonthlyAverageAtPosition(i);

                        Coordinates.Add(new Coordinate(x, y));
                    }


                    break;
                default:
                    //Should not happen unless new trendline types arrive.
                    throw new NotImplementedException("Trendline type not implemented.");
            }
        }

        private void CreateDatalabel()
        {
            if(_trendline.DisplayEquation==false && _trendline.DisplayRSquaredValue==false)
            {
                return;
            }
            //Display the label for the trendline with equation and R² value.
            var lbl = _trendline.Label;
            var coord = RenderCoordinates;
            var x = _svgChart.Plotarea.Rectangle.Left + coord[coord.Length - 2];
            var y = _svgChart.Plotarea.Rectangle.Top + coord[coord.Length - 1];
            double width = 0, height = 0;

            if (_trendline.Label.Layout.HasLayout)
            {
                var mlRect = GetRectFromManualLayout(_svgChart, _trendline.Label.Layout);
                x += mlRect.Left;
                y += mlRect.Top;
                if (lbl.Layout.ManualLayout.Width.HasValue && lbl.Layout.ManualLayout.Height.HasValue)
                {
                    width = mlRect.Width;
                    height = mlRect.Height;
                }
            }

            if (width > 0 && height > 0)
            {
                DataLabel = new SvgTextBox(_svgChart, _svgChart.ChartArea.Rectangle.Bounds, x, y, width, height);
                DataLabel.TextBody.AutoSize = false;
            }
            else
            {
                DataLabel = new SvgTextBox(_svgChart, _svgChart.ChartArea.Rectangle.Bounds, _svgChart.ChartArea.Rectangle.Bounds);
                if (x > 0)
                {
                    DataLabel.Left = x;
                }
                if (y > 0)
                {
                    DataLabel.Top = y;
                }
                DataLabel.TextBody.AutoSize = true;
            }
            //DataLabel.ImportTextBody(lbl.TextBody, true, OfficeOpenXml.Style.ExcelHorizontalAlignment.Center);
            var labelText = "";
            if (_trendline.DisplayEquation)
            {
                labelText += Formula;
            }
            if (_trendline.DisplayRSquaredValue)
            {
                if (labelText.Length > 0)
                {
                    labelText += Environment.NewLine;
                }
                labelText += RSquare;
            }
            DataLabel.ImportParagraph(lbl.TextBody.Paragraphs[0], 0, labelText);
            //DataLabel.AddText(0, labelText);
            //DataLabel.TextBody.Paragraphs[0].AddText(labelText, 0);
            DataLabel.LeftMargin = DataLabel.RightMargin = 4;
            DataLabel.TopMargin = DataLabel.BottomMargin = 2;
            
            //Set datalabel position.
            if(DataLabel.Left - (DataLabel.Width + 5) > _svgChart.Bounds.Right)
            {
                DataLabel.Left = _svgChart.Bounds.Right - DataLabel.Width;
            }
            else
            {
                DataLabel.Left -= (DataLabel.Width + 5);
            }

            if(DataLabel.Left<0)
            {
                DataLabel.Left = 0;
            }

            if(DataLabel.Top - DataLabel.Height / 2 > _svgChart.Bounds.Bottom)
            {
                DataLabel.Top = _svgChart.Bounds.Bottom - DataLabel.Height / 2;
            }
            else
            {
                DataLabel.Top -= DataLabel.Height / 2;
            }

            if (DataLabel.Top < 0)
            {
                DataLabel.Top = 0;
            }


            DataLabel.Rectangle.SetDrawingPropertiesFill(_trendline.Label.Fill, _svgChart.Chart.StyleManager.Style.TrendlineLabel.FillReference.Color);
            DataLabel.Rectangle.SetDrawingPropertiesBorder(_trendline.Label.Border, _svgChart.Chart.StyleManager.Style.TrendlineLabel.BorderReference.Color, true, _trendline.Label.Border.Width);
            DataLabel.Rectangle.SetDrawingPropertiesEffects(_trendline.Label.Effect);
        }

        private void CalculateLinear()
        {
            var n = _xSerie.Count;
            var sumX = n * (n + 1) / 2;
            var sumY = _ySerie.Sum(y => y);
            var sumX2 = n * (n + 1) * (2 * n + 1) / 6;
            var sumXY = 0D;

            double slope, intercept;
            if (double.IsNaN(_trendline.Intercept))
            {
                for (int i = 0; i < _ySerie.Length; i++)
                {
                    sumXY += _ySerie[i] * (i + 1);
                }

                //Slope
                slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
                //Intercept 
                intercept = (sumY - slope * sumX) / n;
            }
            else
            {
                intercept = _trendline.Intercept;
                for (int i = 0; i < _ySerie.Length; i++)
                {
                    sumXY += (_ySerie[i] - intercept) * (i + 1);
                }
                slope = sumXY / sumX2;
            }

            //var r2 = Math.Pow(Pearson.PearsonImpl(_ySerie.Cast<double>(), GetLinearSerie(slope, intercept)), 2);
            var r2 = CalculateRSquared(x => slope * x + intercept, _ySerie, _trendline.Intercept);
            Coefficients = [slope, intercept];
            Formula = $"y={slope:G5}x{(GetValueAndSignSuppressZero(intercept))}";
            RSquare = $"R²={r2:N4}";
        }

        private string GetValueAndSignSuppressZero(double value)
        {
            if(value>0)
            {
                return $"+{value.ToString("G5")}";
            }
            else if(value < 0)
            {
                return $"-{value.ToString("G5")}";
            }
            return "";
        }

        private void CalculateExponential()
        {
            var n = _xSerie.Count;
            var sumX = n * (n + 1) / 2;
            var sumLnY = _ySerie.Sum(y => Math.Log(y));
            var sumX2 = n * (n + 1) * (2 * n + 1) / 6;

            var sumXLnY = 0D;

            //Slope
            double slope, intercept;
            if(double.IsNaN(_trendline.Intercept))
            {
                for (var i = 0; i < _ySerie.Length; i++)
                {
                    sumXLnY += Math.Log(_ySerie[i]) * (i + 1);
                }

                slope = (n * sumXLnY - sumX * sumLnY) / (n * sumX2 - sumX * sumX);
                //Intercept 
                intercept = Math.Pow(Math.E, (sumLnY - slope * sumX) / n);
            }
            else
            {
                intercept = _trendline.Intercept;
                var logIntercept = Math.Log(intercept);
                for (var i = 0; i < _ySerie.Length; i++)
                {
                    sumXLnY += (Math.Log(_ySerie[i]) - logIntercept) * (i + 1);
                }

                slope = sumXLnY / sumX2;                
            }


            var r2 = Math.Pow(Pearson.PearsonImpl(_ySerie.Cast<double>(), GetExponentialSerie(slope, intercept)), 2);
            Coefficients = [slope, intercept];
            Formula = $"y={intercept:G5}|ss:e{slope:G3}|";
            RSquare = $"R²={r2:N4}";
        }
        private void CalculateLogarithmic()
        {
            var n = _xSerie.Count;
            var logSerie =  _xSerie.Select(x => Math.Log(_xSerie.IndexOf(x) + 1)).ToList();
            var sumLnX = logSerie.Sum(x => x);
            var sumLnX2 = logSerie.Sum(x => x * x);
            var sumY = _ySerie.Sum(x => ConvertUtil.GetValueDouble(x));

            var sumLnXY = 0D;// _ySerie.Sum(x => ConvertUtil.GetValueDouble(x) * Math.Log(_ySerie.IndexOf(x) + 1));
            for (int i = 0; i < _ySerie.Length; i++)
            {
                sumLnXY += _ySerie[i] * logSerie[i];
            }



            //Slope
            var slope = (n * sumLnXY - sumLnX * sumY) / (n * sumLnX2 - sumLnX * sumLnX);
            ////Intercept 
            var intercept = (sumY - slope * sumLnX) / n;

            Coefficients = [slope, intercept];

            var r2 = CalculateRSquared(x => slope * Math.Log(x) + intercept, _ySerie, _trendline.Intercept);
            Formula = $"y={slope:G5}ln(x)+{intercept:G5}";
            RSquare = $"R²={r2:N4}";
        }
        public void CalculatePolynomial()
        {

            /*
              Σy    = c·n   + b·Σx  + a·Σx²
              Σxy   = c·Σx  + b·Σx² + a·Σx³
              Σx²y  = c·Σx² + b·Σx³ + a·Σx⁴
              Σx³y  = c·Σx³ + b·Σx⁴ + a·Σx⁵
              ...
            */

            var isForced = double.IsNaN(_trendline.Intercept) == false;
            int n = _ySerie.Length;            
            var order = Math.Min((int)_trendline.Order, n-1);
            int coeffCount = order + (isForced ? 0 : 1);

            // Step 1: Build sums
            double[] sumX = new double[2 * order + 1];
            double[] sumXY = new double[order + 1];

            if (isForced)
            {
                //Todo:Add intercept to the formula and adjust the y values accordingly
                var intercept = _trendline.Intercept;
                double[] yAdj = new double[n];
                for (int i = 0; i < n; i++)
                {
                    yAdj[i] = _ySerie[i] - intercept;
                }

                for (int i = 0; i < n; i++)
                {
                    double xPow = 1.0;
                    for (int k = 0; k <= 2 * order; k++)
                    {
                        sumX[k] += xPow;
                        if (k <= order)
                            sumXY[k] += xPow * yAdj[i];
                        xPow *= i + 1;
                    }
                }
            }
            else
            {
                for (int i = 0; i < n; i++)
                {
                    double xPow = 1.0;
                    for (int k = 0; k <= 2 * order; k++)
                    {
                        sumX[k] += xPow;
                        if (k <= order)
                            sumXY[k] += xPow * _ySerie[i];
                        xPow *= i + 1;
                    }
                }
            }

            // Step 2: Build augmented matrix
            double[,] matrix = new double[coeffCount, coeffCount + 1];

            int offset = isForced ? 2 : 0;
            for (int row = 0; row < coeffCount; row++)
            {
                for (int col = 0; col < coeffCount; col++)
                    matrix[row, col] = sumX[row + col + offset];
                matrix[row, coeffCount] = sumXY[row + (offset / 2)];
            }

            // Step 3: Gaussian elimination with partial pivoting
            for (int i = 0; i < coeffCount; i++)
            {
                // Find best pivot
                int maxRow = i;
                for (int k = i + 1; k < coeffCount; k++)
                {
                    if (Math.Abs(matrix[k, i]) > Math.Abs(matrix[maxRow, i]))
                        maxRow = k;
                }

                // Swap rows
                for (int k = 0; k <= coeffCount; k++)
                {
                    double temp = matrix[i, k];
                    matrix[i, k] = matrix[maxRow, k];
                    matrix[maxRow, k] = temp;
                }

                // Divide pivot row
                double pivot = matrix[i, i];
                for (int k = 0; k <= coeffCount; k++)
                    matrix[i, k] /= pivot;

                // Eliminate column from other rows
                for (int j = 0; j < coeffCount; j++)
                {
                    if (j != i)
                    {
                        double factor = matrix[j, i];
                        for (int k = 0; k <= coeffCount; k++)
                            matrix[j, k] -= factor * matrix[i, k];
                    }
                }
            }

            // Step 4: Extract coefficients
            Coefficients = new double[order+1];
            if (isForced)
            {
                Coefficients[0] = _trendline.Intercept;
                for (int i = 0; i < coeffCount; i++)
                {
                    Coefficients[i+1] = matrix[i, coeffCount];
                }
            }
            else
            {
                for (int i = 0; i < coeffCount; i++)
                {
                    Coefficients[i] = matrix[i, coeffCount];
                }
            }
            Formula = "y=" + GetPolynormFormula();
            var r2 = CalculateRSquared(x => PredictLinear(x), _ySerie, _trendline.Intercept);
            RSquare = $"R²={r2:N4}";
        }

        private void CalculatePower()
        {
            var n = _xSerie.Count;
            var lnSerie = _xSerie.Select((x, i) => Math.Log(i + 1));
            var sumLnX = lnSerie.Sum(x => x);
            var sumLnX2 = lnSerie.Sum(x => x*x);
            var sumLnY = _ySerie.Sum(y => Math.Log(y));
            //double sumLnXLnY = _ySerie.Sum(y => Math.Log(ConvertUtil.GetValueDouble(y)) * Math.Log(_ySerie.IndexOf(y) + 1));
            double sumLnXLnY = 0;
            for (int i=0;i < _ySerie.Length;i++)
            {
                sumLnXLnY += Math.Log(_ySerie[i]) * Math.Log(i + 1);
            }

            var slope = (n * sumLnXLnY - sumLnX * sumLnY) / (n*sumLnX2 - sumLnX * sumLnX);
            var intercept = Math.Pow(Math.E, (sumLnY - slope  * sumLnX) / n);
            Coefficients = [slope, intercept];

            Formula = $"y={intercept:G5}x|ss:{slope:G3}|";
            var ylogSerie = _ySerie.Select(y => Math.Log(y)).ToArray();
            var r2 = CalculateRSquaredPearson(x => intercept * Math.Pow(x, slope), _ySerie);
            RSquare = $"R²={r2:N4}";
        }

        private void CalculateMoveingAverage()
        {
            int n = _ySerie.Length;
            double[] result = new double[n];
            var period = (int)(double.IsNaN(_trendline.Period)  || _trendline.Period < 2 ? 2 : _trendline.Period);
            for (int i = period - 1; i < n; i++)
            {
                double sum = 0;
                for (int j = 0; j < period; j++)
                    sum += _ySerie[i - j];
                result[i] = sum / period;
            }

            Coefficients = result;

            Formula = "";
            RSquare = "";
        }

        private string GetPolynormFormula()
        {
            var sb = new StringBuilder(Coefficients[0].ToString("G5"));
            for (var i = 1; i < Coefficients.Length; i++)
            {
                if (Coefficients[i-1]>=0) 
                { 
                    sb.Insert(0, "+");
                }
                else
                {
                    sb.Insert(0, "-");
                }

                if(i < 2)
                {
                    sb.Insert(0, $"{Math.Abs(Coefficients[i]):G5}x");
                }
                else
                {
                    sb.Insert(0, $"{Math.Abs(Coefficients[i]):G5}x|ss:{i}|");
                }                    
            }
            if (Coefficients[Coefficients.Length - 1] < 0)
            {
                sb.Insert(0, "-");
            }

            return sb.ToString();
        }

        // Predict a y value for a given x
        public double PredictLinear(double x)
        {
            double y = 0;
            double xPow = 1.0;
            for (int i = 0; i < Coefficients.Length; i++)
            {
                y += Coefficients[i] * xPow;
                xPow *= x;
            }
            return y;
        }
        // Predict a y value for a given x
        public double PredictPolynomial(double x)
        {
            double y = 0;
            double xPow = 1.0;
            for (int i = 0; i < Coefficients.Length; i++)
            {
                y += Coefficients[i] * xPow;
                xPow *= x;
            }
            return y;
        }
        private IEnumerable<double> GetExponentialSerie(double m, double b)
        {
            var l = new List<double>() { b };
            for (var i = 1; i < _xSerie.Count; i++)
            {
                l.Add(b * Math.Pow(Math.E, m * i));
            }
            return l;
        }
        public double CalculateRSquared(Func<double, double> predictFunc, double[] serieY, double forcedIntercept)
        {
            int n = serieY.Length;
            if (!double.IsNaN(forcedIntercept))
            {
                // Forced intercept: use squared Pearson correlation
                double meanY = serieY.Average();
                double[] predicted = new double[n];
                for (int i = 0; i < n; i++)
                    predicted[i] = predictFunc(i + 1);
                double meanP = predicted.Average();

                double sumYP = 0, sumY2 = 0, sumP2 = 0;
                for (int i = 0; i < n; i++)
                {
                    double dy = serieY[i] - meanY;
                    double dp = predicted[i] - meanP;
                    sumYP += dy * dp;
                    sumY2 += dy * dy;
                    sumP2 += dp * dp;
                }

                double r = sumYP / Math.Sqrt(sumY2 * sumP2);
                return r * r;
            }
            else
            {
                // Standard R²
                double meanY = serieY.Average();
                double ssRes = 0, ssTot = 0;
                for (int i = 0; i < n; i++)
                {
                    double residual = serieY[i] - predictFunc(i + 1);
                    double deviation = serieY[i] - meanY;
                    ssRes += residual * residual;
                    ssTot += deviation * deviation;
                }
                return 1.0 - (ssRes / ssTot);
            }
        }
        public double CalculateRSquaredPearson(Func<double, double> predictFunc, double[] serieY)
        {
            int n = serieY.Length;
            double meanA = serieY.Average();

            double[] predicted = new double[n];
            for (int i = 0; i < n; i++)
                predicted[i] = predictFunc(i + 1);
            double meanP = predicted.Average();

            double sumAP = 0, sumA2 = 0, sumP2 = 0;
            for (int i = 0; i < n; i++)
            {
                double da = serieY[i] - meanA;
                double dp = predicted[i] - meanP;
                sumAP += da * dp;
                sumA2 += da * da;
                sumP2 += dp * dp;
            }

            double r = sumAP / Math.Sqrt(sumA2 * sumP2);
            return r * r;
        }
        public double[] Coefficients {get;set;}
        public string Formula { get; set; }
        public string RSquare { get; set; }
        public SvgTextBox DataLabel { get; set; }
        public override string ToString()
        {
            return _trendline.Type + "," + Formula + "," + RSquare;
        }
        internal void CreateRenderCoordinatesAndDatalabel()
        {
            CreateRenderCoordinates();
            CreateDatalabel();
        }
        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var pathItem = new SvgRenderPathItem(_svgChart, _svgChart.Plotarea.Rectangle.Bounds);
            pathItem.Commands.Add(new EPPlusImageRenderer.PathCommands(PathCommandType.Move, pathItem, RenderCoordinates));
            pathItem.FillColor = "none";
            pathItem.SetDrawingPropertiesBorder(_trendline.Border, _svgChart.Chart.StyleManager.Style.Trendline.BorderReference.Color, true, _trendline.Border.Width);
            pathItem.SetDrawingPropertiesEffects(_trendline.Effect);
            renderItems.Add(pathItem);
        }

        private void CreateCoordinates(Func<double, double> predictPoint)
        {
            var x1 = 0;
            var x2 = _xSerie.Count - 1;

            //We aim for 1 line per point for the trendline.
            var diff = (x2 - x1);
            var inc = diff / GetXInc(_svgChart.Bounds.Width);
            double y;
            for (double d = x1; d < x2; d += inc)
            {
                var x = (x2) * (d - x1) / diff;
                y = predictPoint(x + 1);
                Coordinates.Add(new Coordinate(x,y));
            }

            y = predictPoint(x2 + 1);
            Coordinates.Add(new Coordinate(x2, y));
        }
        private void CreateRenderCoordinates()
        {
            var isBar = _chartType.IsTypeBar();
            var isLine = _chartType.IsTypeLine();
            SvgChartAxis catAxis, valAxis;
            if(isBar)
            {
                valAxis = _useSecondaryAxis ? _svgChart.SecondHorizontalAxis : _svgChart.HorizontalAxis;
                catAxis = _useSecondaryAxis ? _svgChart.SecondVerticalAxis : _svgChart.VerticalAxis;
            }
            else
            {
                catAxis = _useSecondaryAxis ? _svgChart.SecondHorizontalAxis : _svgChart.HorizontalAxis;
                valAxis = _useSecondaryAxis ? _svgChart.SecondVerticalAxis : _svgChart.VerticalAxis;
            }

            var pa = _svgChart.Plotarea;
            var coordinates=new List<double>();
            for (var i=0;i<Coordinates.Count;i++)
            {
                if(isLine)
                {
                    coordinates.Add(catAxis.GetPositionInPlotarea(Coordinates[i].X));
                    coordinates.Add(valAxis.GetPositionInPlotarea(Coordinates[i].Y));
                }
                else
                {
                    var count = (_xSerie.Count > _ySerie.Length ? _xSerie.Count: _ySerie.Length);
                    var ct = (ExcelBarChart) _chartType;
                    var yWidth = (isBar ? _svgChart.Plotarea.Rectangle.Height : _svgChart.Plotarea.Rectangle.Width);
                    var slotSize = valAxis.Values.Count;
                    var gapPercent = ct.GapWidth / 100D;     // Gap width between bars/columns in percent
                    var overlapPercent = ct.Overlap / 100D;  // Overlap  between bars/columns in percent            
                    var slotWidth = yWidth / slotSize;
                    var clusterWidth = slotWidth * 100 / (100 + ct.GapWidth);
                    var step = 1 - overlapPercent;
                    var barWidth = slotWidth / (1 + (count - 1) * step + gapPercent);
                    var halfGap = (barWidth * gapPercent) / 2;
                    if (isBar)
                    {

                        coordinates.Add(valAxis.GetPositionInPlotarea(Coordinates[i].Y));
                        coordinates.Add(catAxis.GetPositionInPlotarea(Coordinates[Coordinates.Count - 1].X - Coordinates[i].X) + halfGap + (_serieCount - _seriePos - 1) * barWidth * step);
                    }
                    else
                    {
                        coordinates.Add(catAxis.GetPositionInPlotarea(Coordinates[i].X));
                        coordinates.Add(valAxis.GetPositionInPlotarea(Coordinates[i].Y));
                    }
                }
            }
            RenderCoordinates = coordinates.ToArray();
        }

        //Get the incremental x value for the trendline points based on the distance between the start and end point of the trendline. The goal is to have approximately 3 point per data point in the trendline.
        private double GetXInc(double n)
        {
            int k = (int)Math.Round(Math.Log(n)/ Math.Log(3) - 1);
            k = Math.Max(k, 0);
            return n / Math.Pow(2, k);
        }
        internal List<Coordinate> Coordinates { get; set; } = new List<Coordinate>();
        internal double[] RenderCoordinates { get; set; }
        List<double> _ma = null;
        private double GetMonthlyAverageAtPosition(double x)
        {
            if (_ma == null)
            {
                CalcMa();
            }

            int ix = (int)(x - _trendline.Period + 1);
            return _ma[ix];
        }

        private void CalcMa()
        {
            _ma= new List<double>();
            double sum = 0;
            for (int i=0;i < _ySerie.Length;i++)
            {
                sum += _ySerie[i];
                if (i >= _trendline.Period-1)
                {
                    
                    _ma.Add(sum / (i+1));
                }
            }
        }

        private double GetLinearValueAtPosition(int x)
        {
            return Coefficients[1] + Coefficients[0] * x;
        }
    }
}