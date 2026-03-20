using EPPlusImageRenderer.Svg;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Utils.DateUtils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace EPPlus.Export.ImageRenderer.Svg.Chart.Util
{
    internal class CategoryAxisScaleCalculator
    {

        internal static AxisScale CalculateByWidth(ref List<object> values, ITextMeasurer tm, AxisOptions options)
        {
            //var height = options.ChartSize.Bounds.Height;
            var ax = options.Axis;
            var plotAreaWidth = options.ChartSize.Bounds.Width;
            var mf = ax.Font.GetMeasureFont();
            var displayValues = values.Select(x => ax.ToString()).ToList();
            var res = tm.MeasureText(displayValues[0], mf);

            //Get interval for maximum width with vertical text.
            var interval = GetMinUnitVerticalText(displayValues, res.Height, plotAreaWidth);


            //Get max text width when using diagonal text
            var width = mf.Size * Math.Sqrt(2);
            var margin = mf.Size * 0.5;

            if (FitAsVerticalDiagonalText(displayValues, interval, width, margin, plotAreaWidth)) //Check diagonal
            {
                if(FitAsHorizontalText(displayValues, interval, mf, tm, plotAreaWidth)) //Check horizontal
                {
                    return new AxisScale()
                    {
                        MajorInterval = interval,
                        MinorInterval = 1,
                        Min = 1,
                        Max = displayValues.Count,
                        TextOrientation = eTextOrientation.Horizontal
                    };
                }
                else
                {
                    return new AxisScale()
                    {
                        MajorInterval = interval,
                        MinorInterval = 1,
                        Min = 1,
                        Max = displayValues.Count,
                        TextOrientation = eTextOrientation.Diagonal
                    };
                }
            }
            if(interval!=1)
            {
                var removeCount = interval - 1;
                var c = (int)Math.Truncate(values.Count / (double)interval);
                for(int i=0;i<=c;i++)
                {
                    for(int j=0;j<removeCount;j++)
                    {
                        if (i + 1 < values.Count)
                        {
                            values.RemoveAt(i + 1);
                        }
                    }
                }
            }
            return new AxisScale()
            {
                MajorInterval = interval,
                MinorInterval = 1,
                Min = 1,
                Max = displayValues.Count,
                TextOrientation = eTextOrientation.Vertical
            };

        }

        private static bool FitAsHorizontalText(List<string> displayValues, int interval, MeasurementFont mf, ITextMeasurer tm, double plotAreaWidth)
        {
            var margin = mf.Size * 0.3;
            var width = tm.MeasureText(displayValues[0], mf).Width + margin;
            var pos = interval;
            while (pos < displayValues.Count && width < plotAreaWidth)
            {
                pos += interval;
                width = tm.MeasureText(displayValues[pos], mf).Width + margin;
                if(width > plotAreaWidth) return false;
            }
            return width <= plotAreaWidth;
        }

        private static bool FitAsVerticalDiagonalText(List<string> displayValues, int interval, double textWidth, double margin, double plotAreaWidth)
        {
            var items = Math.Truncate(displayValues.Count * 1D / interval);
            return items * textWidth + (items - 1) * margin < plotAreaWidth;
        }

        private static int GetMinUnitVerticalText(List<string> displayValues, double textHeight, double plotAreaWidth)
        {
            var interval = 1;
            var margin = 0D;
            var items = Math.Truncate((double)displayValues.Count / interval);
            while (items * textHeight + (items - 1) * margin  >= plotAreaWidth)
            {
                interval++;
                items = Math.Truncate((double)displayValues.Count / interval);
            }

            return interval;
        }
    }
}