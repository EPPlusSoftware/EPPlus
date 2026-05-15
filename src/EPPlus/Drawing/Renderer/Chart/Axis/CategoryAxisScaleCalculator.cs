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

        internal static AxisScale CalculateHorizontalAxisByWidth(ref List<object> values, ITextMeasurer tm, AxisOptions options)
        {
            var ax = options.Axis;
            var plotAreaWidth = options.ChartSize.Bounds.Width;
            var mf = ax.Font.GetMeasureFont();
            List<object> displayValues = GetUniqueValues(values).Select(x=>(object)x.ToString()).ToList();
            var uniqeItems = displayValues.Count;
            var res = tm.MeasureText(displayValues[0].ToString(), mf);

            //Get interval for maximum width with vertical text.
            var interval = GetMinUnitVerticalText(displayValues.Count, res.Height, plotAreaWidth);


            //Get max text width when using diagonal text
            var width = mf.Size * Math.Sqrt(2);
            var margin = mf.Size * 0.5;

            if (FitAsVerticalDiagonalText(displayValues.Count, interval, width, margin, plotAreaWidth)) //Check diagonal
            {
                if(FitAsHorizontalText(displayValues, interval, mf, tm, plotAreaWidth)) //Check horizontal
                {
                    return new AxisScale()
                    {
                        MajorInterval = interval,
                        MinorInterval = 1,
                        Min = 1,
                        Max = displayValues.Count,
                        TextOrientation = eTextOrientation.Horizontal,
                        DisplayValues = displayValues
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
                        TextOrientation = eTextOrientation.Diagonal,
                        DisplayValues = displayValues
                    };
                }
            }
            if(interval != 1)
            {
                var removeCount = interval - 1;
                var c = (int)Math.Truncate(values.Count / (double)interval);
                for(int i=0;i<=c;i++)
                {
                    for(int j=0;j<removeCount;j++)
                    {
                        if (i + 1 < displayValues.Count)
                        {
                            displayValues.RemoveAt(i + 1);
                        }
                    }
                }
            }

            return new AxisScale()
            {
                MajorInterval = interval,
                MinorInterval = 1,
                Min = 1,
                Max = uniqeItems,
                TextOrientation = eTextOrientation.Vertical,
                DisplayValues = displayValues
            };

        }
        internal static AxisScale CalculateVerticalAxisByHeight(ref List<object> values, ITextMeasurer tm, AxisOptions options)
        {
            var ax = options.Axis;
            var plotAreaHeight = options.ChartSize.Bounds.Height;
            var mf = ax.Font.GetMeasureFont();
            List<object> displayValues = GetUniqueValues(values).Select(x => (object)x.ToString()).ToList();
            var uniqeItems = displayValues.Count;
            var res = tm.MeasureText(displayValues[0].ToString(), mf);
            int interval = 1;
            while (FitAsVerticalDiagonalText(displayValues.Count, interval, res.Height, 1.2D, plotAreaHeight)==false) //Check horizontal
            {
                interval++;
            }
            if (interval != 1)
            {
                var removeCount = interval-1;
                var c = (int)Math.Truncate(displayValues.Count / (double)interval);
                for (int i = 0; i <= c; i++)
                {
                    for (int j = 0; j < removeCount; j++)
                    {
                        if (i + 1 < displayValues.Count)
                        {
                            displayValues.RemoveAt(i + 1);
                        }
                    }
                }
            }

            return new AxisScale()
            {
                MajorInterval = interval,
                MinorInterval = 1,
                Min = 1,
                Max = uniqeItems,
                TextOrientation = eTextOrientation.Horizontal,
                DisplayValues = displayValues
            };
        }
        internal static List<object> GetUniqueValues(List<object> values)
        {
            var ret = new List<object>();
            var hs = new HashSet<string>();
            foreach(var v in values)
            {
                if(v is string[])
                {
                    var s = (string[])v;
                    var key = s[0] + s[1];
                    if(hs.Add(key))
                    {
                        ret.Add(s[0]);
                    }
                }
                else
                {
                    ret.Add(v);
                }
            }
            return ret;
        }

        private static bool FitAsHorizontalText(List<object> displayValues, int interval, MeasurementFont mf, ITextMeasurer tm, double plotAreaWidth)
        {
            var margin = mf.Size * 0.3;
            var width = tm.MeasureText(displayValues[0].ToString(), mf).Width + margin;
            var pos = interval;
            while (pos < displayValues.Count && width < plotAreaWidth)
            {
                width = tm.MeasureText(displayValues[pos].ToString(), mf).Width + margin;
                if(width > plotAreaWidth) return false;
                pos += interval;
            }
            return width <= plotAreaWidth;
        }

        private static bool FitAsVerticalDiagonalText(int itemCount, int interval, double textWidth, double margin, double plotAreaWidthHeight)
        {
            var items = Math.Truncate(itemCount * 1D / interval);
            return items * textWidth + (items - 1) * margin < plotAreaWidthHeight;
        }

        private static int GetMinUnitVerticalText(int itemCount, double textHeight, double plotAreaWidth)
        {
            var interval = 1;
            var margin = 0D;
            var items = Math.Truncate((double)itemCount / interval);
            while (items * textHeight + (items - 1) * margin  >= plotAreaWidth)
            {
                interval++;
                items = Math.Truncate((double)itemCount / interval);
            }

            return interval;
        }
    }
}