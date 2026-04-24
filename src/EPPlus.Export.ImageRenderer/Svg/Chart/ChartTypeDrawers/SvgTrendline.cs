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
using EPPlusImageRenderer;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Statistical;
using OfficeOpenXml.Packaging.Ionic;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal class SvgTrendline
    {
        private ExcelChartTrendline _trendline;
        private double[] _ySerie;
        private List<object> _xSerie;

        public SvgTrendline(ExcelChartTrendline trendline, List<object> xSerie, List<object> ySerie)
        {
            _trendline = trendline;
            _xSerie = xSerie;
            _ySerie = ySerie.Select(y => ConvertUtil.GetValueDouble(y)).ToArray();
            switch(trendline.Type) 
            {
                case eTrendLine.Linear:
                    CalculateLinear();
                    break;
                case eTrendLine.Exponential:
                    CalculateExponential();
                    break;
                case eTrendLine.Logarithmic:
                    CalculateLogarithmic();
                    break;
                case eTrendLine.Polynomial:
                    CalculatePolynomial();
                    break;
                case eTrendLine.Power:
                    CalculatePower();
                    break;
                case eTrendLine.MovingAverage:
                    CalculateMoveingAverage();
                    break;
                default:
                    throw new NotImplementedException("Trendline type not implemented.");
            }
        }

        private double[] GetXValues(List<object> xSerie)
        {
            var ret = new double[xSerie.Count];
            for(int i=0;i<xSerie.Count;i++)
            {
                ret[i] = i-1;
            }
            return ret;
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
            var r2 = CalculateRSquared(x => slope * x + intercept, _ySerie);
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

                slope = (n * sumXLnY - _trendline.Intercept * sumLnY) / (n * sumX2 - sumX * sumX);
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

            var r2 = CalculateRSquared(x => slope * Math.Log(x) + intercept, _ySerie);
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


            int n = _ySerie.Length;            
            var order = Math.Min((int)_trendline.Order, n-1);
            int coeffCount = order + 1;

            // Step 1: Build sums
            double[] sumX = new double[2 * order + 1];
            double[] sumXY = new double[order + 1];

            if (double.IsNaN(_trendline.Intercept))
            {
                var intercept = _trendline.Intercept;
                double[] yAdj = new double[n];
                for (int i = 0; i < n; i++)
                {
                    yAdj[i] = _ySerie[i] - intercept;
                }
            }
            else
            {
                //Todo:Add intercept to the formula and adjust the y values accordingly
            }

            for (int i = 0; i < n; i++)
            {
                double xPow = 1.0;
                for (int k = 0; k <= 2 * order; k++)
                {
                    sumX[k] += xPow;
                    if (k <= order)
                        sumXY[k] += xPow * ConvertUtil.GetValueDouble(_ySerie[i]);
                    xPow *= i + 1;
                }
            }

            // Step 2: Build augmented matrix
            double[,] matrix = new double[coeffCount, coeffCount + 1];

            for (int row = 0; row < coeffCount; row++)
            {
                for (int col = 0; col < coeffCount; col++)
                    matrix[row, col] = sumX[row + col];
                matrix[row, coeffCount] = sumXY[row];
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
            // coefficients[0] = constant, [1] = x term, [2] = x² term, etc.
            Coefficients = new double[coeffCount];
            for (int i = 0; i < coeffCount; i++)
            {
                Coefficients[i] = matrix[i, coeffCount];
            }
            Formula = "y=" + GetPolynormFormula();
            RSquare = "R²=" + CalculateRSquared(x => PredictPolynomial(x), _ySerie);
        }

        private void CalculatePower()
        {
            var n = _xSerie.Count;
            var lnSerie = _xSerie.Select(x => Math.Log(_xSerie.IndexOf(x) + 1));
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

            Formula = $"y={intercept:G5}x|ss:{slope:G5}";
            var ylogSerie = _ySerie.Select(y => Math.Log(y)).ToArray();
            RSquare = "R²=" + CalculateRSquared(x => slope * x + Math.Log(intercept), ylogSerie);
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

                sb.Insert(0,$"{Math.Abs(Coefficients[i]):G5}x|ss:{i+1}|");
            }
            if (Coefficients[Coefficients.Length - 1] < 0)
            {
                sb.Insert(0, "-");
            }

            return sb.ToString();
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
        private IEnumerable<double> GetLinearSerie(double m, double b)
        {
            var l = new List<double>() { b };
            for (var i = 1; i < _ySerie.Length; i++)
            {
                l.Add(b + m * i);
            }
            return l;
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
        private IEnumerable<double> GetPowerSerie(double m, double b)
        {
            var l = new List<double>() { b };
            for (var i = 1; i < _xSerie.Count; i++)
            {
                l.Add(b * Math.Pow(Math.E, m * i));
            }
            return l;
        }
        //private IEnumerable<double> GetLogarithmicSerie(double m, double b)
        //{
        //    var l = new List<double>() { b };
        //    for (var i = 1; i < _ySerie.Count; i++)
        //    {
        //        l.Add(m * Math.Log(i + 1) + b);
        //    }
        //    return l;
        //}
        public double CalculateRSquared(Func<double, double> predictFunc, double[] serieY)
        {
            int n = serieY.Length;

            double avgY = serieY.Sum(y => y) / n;

            double ssRes = 0;
            double ssTot = 0;

            for (int i = 0; i < n; i++)
            {
                var v = serieY[i];
                double residual = v - predictFunc(i + 1);
                double deviation = v - avgY;

                ssRes += residual * residual;
                ssTot += deviation * deviation;
            }

            return 1.0 - (ssRes / ssTot);
        }
        public double[] Coefficients {get;set;}
        public string Formula { get; set; }
        public string RSquare { get; set; }

        public override string ToString()
        {
            return _trendline.Type + "," + Formula + "," + RSquare;
        }
    }
}