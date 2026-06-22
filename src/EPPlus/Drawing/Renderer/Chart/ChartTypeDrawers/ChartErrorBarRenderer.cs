using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;
using static OfficeOpenXml.ExcelErrorValue;

namespace OfficeOpenXml.Drawing.Renderer.Chart.ChartTypeDrawers
{
    internal class ChartErrorBarRenderer : ChartDrawingObject
    {
        internal ExcelChartErrorBars _errorbars;
        private double[] _ySerie;
        private List<double> _xSerie;
        private ExcelChart _chartType;
        
        internal List<double[]> Values { get; set; } = new List<double[]>();
        public ChartErrorBarRenderer(ChartRenderer svgChart, ExcelChartErrorBars errorbars, List<object> xSerie, List<object> ySerie, ExcelChart chartType, int seriePos) : base(svgChart)
        {
            _chartType = chartType;
            _errorbars = errorbars;
            _xSerie = GetXSerie(xSerie);
            _ySerie = ySerie.Select(y => ConvertUtil.GetValueDouble(y)).ToArray();
            //_useSecondaryAxis = chartType.UseSecondaryAxis;
            //_serieCount = _chartType.Series.Count;
            //_seriePos = seriePos;
            int n = xSerie.Count;

            switch (errorbars.ValueType)
            {
                case eErrorValueType.StandardError:
                case eErrorValueType.StandardDeviation:
                    if (n > 1)
                    {
                        double avg = _ySerie.Average();
                        double sumSquaredDeviations = _ySerie.Sum(v => (v - avg) * (v - avg));
                        double sampleStdDev = Math.Sqrt(sumSquaredDeviations / (n - 1)); // sample std dev (n-1)
                        for (int i=0;i < _xSerie.Count;i++)
                        {
                            if (i >= _ySerie.Length) break;

                            if(errorbars.ValueType == eErrorValueType.StandardError)
                            {
                                double se = sampleStdDev / Math.Sqrt(n);
                                double y = _ySerie[i];
                                Values.Add(new double[] { y - se, y, y + se });
                            }
                            else
                            {
                                var mult = errorbars.Value ?? 1D;
                                Values.Add(new double[] { avg - sampleStdDev, avg, avg + sampleStdDev });
                            }
                        }
                    }
                    break;
                case eErrorValueType.Percentage:
                    var percent = (errorbars.Value ?? 0D) / 100D;
                    for (int i = 0; i < _xSerie.Count; i++)
                    {
                        double y = _ySerie[i];
                        Values.Add(new double[] { y * (1 - percent), y, y * (1 + percent) });
                    }
                    break;
                case eErrorValueType.FixedValue:
                    var fixedValue = errorbars.Value ?? 0D;
                    for (int i = 0; i < _xSerie.Count; i++)
                    {
                        double y = _ySerie[i];
                        Values.Add(new double[] { y - fixedValue, y, y + fixedValue });
                    }

                    break;
                case eErrorValueType.Custom:
                    var minusList = errorbars.Minus.GetValuesList(_chartType.WorkSheet.Workbook);
                    var plusList = errorbars.Plus.GetValuesList(_chartType.WorkSheet.Workbook);
                    for (int i = 0; i < _xSerie.Count; i++)
                    {
                        double y = _ySerie[i];
                        double minus = GetCustomValue(minusList, i);
                        double plus = GetCustomValue(plusList, i);
                        Values.Add(new double[] { y - minus, y, y + plus });
                    }

                    break;
            }
        }
        public double GetCustomValue(List<double> l, int i)
        {
            if(l.Count==0)
            {
                return l[0];
            }
            else if(i<l.Count)
            {
                return l[i];
            }
            return 0D;
        }

        internal List<RenderItem> GetErrorBarRenderItem(int index, ChartAxisRenderer xAxis, ChartAxisRenderer yAxis, double x, double y, double xPos, double yPos)
        {
            //var path = new PathRenderItem(ChartRenderer.Bounds);
            var l = new List<RenderItem>();
            double topValue=0, bottomValue=0;
            if (_errorbars.BarType == eErrorBarType.Plus || _errorbars.BarType == eErrorBarType.Both)
            {
                topValue = Values[index][2];
            }
            if(_errorbars.BarType == eErrorBarType.Minus || _errorbars.BarType == eErrorBarType.Both)
            {
                bottomValue = Values[index][0];
            }

            if (_errorbars.Direction == eErrorBarDirection.X)
            {
                var rightPos = xAxis.GetPositionInPlotarea(topValue);
                var centerPos = yAxis.GetPositionInPlotarea(Values[index][1]);
                var leftPos = xAxis.GetPositionInPlotarea(bottomValue);

                //Bottom line
                //path.Commands.Add(new PathCommands(PathCommandType.Move, leftPos, yPos, centerPos, yPos));
                var bl = new LineRenderItem(ChartRenderer.Bounds)
                {
                    X1 = leftPos,
                    Y1 = yPos,
                    X2 = centerPos,
                    Y2 = yPos
                };
                //Top line
                //path.Commands.Add(new PathCommands(PathCommandType.Move, centerPos, yPos, rightPos, yPos));
                var tl = new LineRenderItem(ChartRenderer.Bounds)
                {
                    X1 = centerPos,
                    Y1 = yPos,
                    X2 = rightPos,
                    Y2 = yPos
                };
                l.Add(bl);
                l.Add(tl);
                if (_errorbars.NoEndCap == false)
                {
                    //Bottom cap
                    //path.Commands.Add(new PathCommands(PathCommandType.Move, leftPos, yPos - 3, leftPos, yPos + 3));
                    var bc = new LineRenderItem(ChartRenderer.Bounds)
                    {
                        X1 = leftPos,
                        Y1 = yPos - 3,
                        X2 = leftPos,
                        Y2 = yPos + 3
                    };
                    //Top cap
                    //path.Commands.Add(new PathCommands(PathCommandType.Move, rightPos, yPos - 3, rightPos, yPos + 3));
                    var tc = new LineRenderItem(ChartRenderer.Bounds)
                    {
                        X1 = rightPos,
                        Y1 = yPos - 3,
                        X2 = rightPos,
                        Y2 = yPos + 3
                    };
                    l.Add(bc);
                    l.Add(tc);
                }
            }
            else
            {
                var bottomPos = yAxis.GetPositionInPlotarea(bottomValue);
                var centerPos = yAxis.GetPositionInPlotarea(Values[index][1]);
                var topPos = yAxis.GetPositionInPlotarea(topValue);

                //Bottom line
                //path.Commands.Add(new PathCommands(PathCommandType.Move, xPos, bottomPos, xPos, centerPos));
                var bl = new LineRenderItem(ChartRenderer.Bounds)
                {
                    X1 = xPos,
                    Y1 = bottomPos,
                    X2 = xPos,
                    Y2 = centerPos
                };
                //Top line
                //path.Commands.Add(new PathCommands(PathCommandType.Move, xPos, topPos, xPos, centerPos));
                var tl = new LineRenderItem(ChartRenderer.Bounds)
                {
                    X1 = xPos,
                    Y1 = topPos,
                    X2 = xPos,
                    Y2 = centerPos
                };
                l.Add(bl);
                l.Add(tl);
                if (_errorbars.NoEndCap == false)
                {
                    //Bottom cap
                    //path.Commands.Add(new PathCommands(PathCommandType.Move, xPos-3, bottomPos, xPos+3, bottomPos));
                    var bc = new LineRenderItem(ChartRenderer.Bounds)
                    {
                        X1 = xPos - 3,
                        Y1 = bottomPos,
                        X2 = xPos + 3,
                        Y2 = bottomPos
                    };

                    //Top cap
                    //path.Commands.Add(new PathCommands(PathCommandType.Move, xPos - 3, topPos, xPos + 3, topPos));
                    var tc = new LineRenderItem(ChartRenderer.Bounds)
                    {
                        X1 = xPos - 3,
                        Y1 = topPos,
                        X2 = xPos + 3,
                        Y2 = topPos
                    };
                    l.Add(bc);
                    l.Add(tc);
                }
            }
            foreach (var ri in l)
            {
                if (_errorbars.Border.LineElement == null)
                {
                    ri.SetDrawingPropertiesBorder(ChartRenderer.Theme, ChartRenderer.Chart.StyleManager.Style?.ErrorBar.Border, ChartRenderer.Chart.StyleManager.Style?.ErrorBar.BorderReference.Color, true, 0.75);
                }
                else
                {
                    ri.SetDrawingPropertiesBorder(ChartRenderer.Theme, _errorbars.Border, ChartRenderer.Chart.StyleManager.Style?.ErrorBar.BorderReference.Color, _errorbars.Border.Fill.Style != eFillStyle.NoFill, 0.75);
                }
                ri.SetDrawingPropertiesEffects(ChartRenderer.Theme, _errorbars.Effect);
            }
            return l;
        }
    }
}
