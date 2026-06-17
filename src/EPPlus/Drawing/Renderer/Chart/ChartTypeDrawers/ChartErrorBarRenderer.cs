using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
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
                        
                        for(int i=0;i < _xSerie.Count;i++)
                        {
                            if (i >= _ySerie.Length) break;

                            if(errorbars.ValueType == eErrorValueType.StandardError)
                            {
                                double y = _ySerie[i];
                                Values.Add(new double[] { y - sampleStdDev, y + sampleStdDev });
                            }
                            else
                            {
                                var mult = errorbars.Value ?? 1D;
                                sampleStdDev*= mult;
                                Values.Add(new double[] { avg - sampleStdDev, avg + sampleStdDev });
                            }
                        }
                    }
                    break;
                case eErrorValueType.Percentage:
                    var percent = (errorbars.Value ?? 0D) / 100D;
                    for (int i = 0; i < _xSerie.Count; i++)
                    {
                        double y = _ySerie[i];
                        Values.Add(new double[] { y * (1 - percent), y + (1 + percent) });
                    }
                    break;
                case eErrorValueType.FixedValue:
                    var fixedValue = errorbars.Value ?? 0D;
                    for (int i = 0; i < _xSerie.Count; i++)
                    {
                        double y = _ySerie[i];
                        Values.Add(new double[] { y - fixedValue, y + fixedValue });
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
                        Values.Add(new double[] { y - minus, y + plus });
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
        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            throw new NotImplementedException();
        }

        internal RenderItem GetErrorBarRenderItem(int index, ChartAxisRenderer xAxis, ChartAxisRenderer yAxis, double x, double y)
        {
            var path = new PathRenderItem(ChartRenderer.Bounds);
            double addTop=0, addBottom=0;
            if (_errorbars.BarType == eErrorBarType.Plus || _errorbars.BarType == eErrorBarType.Both)
            {
                addTop = Values[index][1];
            }
            if(_errorbars.BarType == eErrorBarType.Minus || _errorbars.BarType == eErrorBarType.Both)
            {
                addBottom = Values[index][0];
            }

            if(_errorbars.Direction == eErrorBarDirection.X)
            {
                path.Commands.Add(new PathCommands(PathCommandType.Move, xAxis.GetPositionInPlotarea(x - addBottom), yAxis.GetPositionInPlotarea(y)));
                path.Commands.Add(new PathCommands(PathCommandType.Line, xAxis.GetPositionInPlotarea(x + addTop), yAxis.GetPositionInPlotarea(y)));
            }
            else
            {
                path.Commands.Add(new PathCommands(PathCommandType.Move, xAxis.GetPositionInPlotarea(x), yAxis.GetPositionInPlotarea(y - addBottom)));
                path.Commands.Add(new PathCommands(PathCommandType.Line, xAxis.GetPositionInPlotarea(x), yAxis.GetPositionInPlotarea(y + addTop)));
            }
            return path;
        }
    }
}
