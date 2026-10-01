using EPPlus.Export.Renderer;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Utils.DateUtils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace EPPlus.Export.ImageRenderer.Svg.Chart.Util
{
    internal class CategoryAxisScaleCalculator
    {

        internal static AxisScale CalculateHorizontalAxisByWidth(ref List<object> values, ITextShaper shaper, float fontSize, AxisOptions options)
        {
            var plotAreaWidth = options.ChartSize.Width;
            List<object> displayValues = GetUniqueValues(values).Select(x => (object)x.ToString()).ToList();
            var uniqeItems = displayValues.Count;

            var textHeight = GetTextHeight(shaper, displayValues[0].ToString(), fontSize);

            //Get interval for maximum width with vertical text.
            var interval = GetMinUnitVerticalText(displayValues.Count, textHeight, plotAreaWidth);

            //Get max text width when using diagonal text
            var width = fontSize * Math.Sqrt(2);
            var margin = fontSize * 0.5;

            if (FitAsVerticalDiagonalText(displayValues.Count, interval, width, margin, plotAreaWidth)) //Check diagonal
            {
                if (FitAsHorizontalText(displayValues, interval, shaper, fontSize, plotAreaWidth)) //Check horizontal
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
            if (interval != 1)
            {
                var removeCount = interval - 1;
                var c = (int)Math.Truncate(values.Count / (double)interval);
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
                TextOrientation = eTextOrientation.Vertical,
                DisplayValues = displayValues
            };
        }
        internal static AxisScale CalculateVerticalAxisByHeight(ref List<object> values, ITextShaper shaper, float fontSize, AxisOptions options)
        {
            var plotAreaHeight = options.ChartSize.Height;
            List<object> displayValues = GetUniqueValues(values).Select(x => (object)x.ToString()).ToList();
            var uniqeItems = displayValues.Count;
            var textHeight = GetTextHeight(shaper, displayValues[0].ToString(), fontSize);
            int interval = 1;
            while (FitAsVerticalDiagonalText(displayValues.Count, interval, textHeight, 1.2D, plotAreaHeight) == false) //Check horizontal
            {
                interval++;
            }
            if (interval != 1)
            {
                var removeCount = interval - 1;
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


        private static bool FitAsHorizontalText(List<object> displayValues, int interval, ITextShaper shaper, float fontSize, double plotAreaWidth)
        {
            var margin = fontSize * 0.3;
            var width = GetTextWidth(shaper, displayValues[0].ToString(), fontSize) + margin;
            var pos = interval;
            while (pos < displayValues.Count && width < plotAreaWidth)
            {
                width = GetTextWidth(shaper, displayValues[pos].ToString(), fontSize) + margin;
                if (width > plotAreaWidth) return false;
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

        internal static List<object> GetUniqueValues(List<object> values)
        {
            var ret = new List<object>();
            var hs = new HashSet<string>();
            foreach (var v in values)
            {
                if (v is object[])
                {
                    var s = (object[])v;
                    var key = s[0].ToString() + s[1].ToString();
                    if (hs.Add(key))
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

        /// <summary>
        /// Width of a single line of text in points. An empty text has no width.
        /// </summary>
        private static double GetTextWidth(ITextShaper shaper, string text, float fontSize)
        {
            if (string.IsNullOrEmpty(text))
                return 0D;
            return shaper.Shape(text).GetWidthInPoints(fontSize);
        }

        /// <summary>
        /// Line height in points, matching the height previously returned by ITextMeasurer.
        /// An empty text has no height.
        /// </summary>
        private static double GetTextHeight(ITextShaper shaper, string text, float fontSize)
        {
            if (string.IsNullOrEmpty(text))
                return 0D;
            return shaper.GetLineHeightInPoints(fontSize);
        }
    }
}