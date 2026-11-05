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
using OfficeOpenXml.Utils.TypeConversion;
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
            var uniqueValues = GetUniqueValues(values, options.IsMulti);
            List<object> displayValues;
            if (options.IsMulti)
            {
                displayValues = uniqueValues.ToList();
            }
            else
            {
                displayValues = uniqueValues.Select(x => (object)x.ToString()).ToList();
            }
            var uniqeItems = displayValues.Count;

            var textHeight = GetTextHeight(shaper, displayValues[0].ToString(), fontSize);

            //Get interval for maximum width with vertical text.
            var interval = GetMinUnitVerticalText(displayValues.Count, textHeight, plotAreaWidth);

            //Get max text width when using diagonal text
            if (options.Axis.TextBody.RotationAdjustedForTextVertical != 0 )
            {
                var orientationCustom = eTextOrientation.Custom;

                if (options.Axis.TextBody.RotationAdjustedForTextVertical == 90)
                {
                    orientationCustom = eTextOrientation.Vertical;
                }
                //else if (options.Axis.TextBody.RotationAdjustedForTextVertical == 45 || (options.Axis.TextBody.RotationAdjustedForTextVertical == -45))
                //{
                //    orientationCustom = eTextOrientation.Diagonal;
                //}


                //var rot = options.Axis.TextBody.Rotation.Value;
                //var width = fontSize * Math.Sqrt(1D/(rot % 90) / 90);
                //var margin = fontSize * 0.5;
                //if (FitAsVerticalDiagonalText(displayValues.Count, interval, width, margin, plotAreaWidth)) //Check diagonal
                //{
                var width = fontSize * Math.Sqrt(2);
                var margin = fontSize * 0.5;
                if (FitAsVerticalDiagonalText(displayValues.Count, interval, width, margin, plotAreaWidth)) //Check diagonal
                {
                    return new AxisScale()
                    {
                        MajorInterval = interval,
                        MinorInterval = 1,
                        Min = 1,
                        Max = displayValues.Count,
                        TextOrientation = orientationCustom,
                        DisplayValues = displayValues
                    };
                }
                else if(FitAsVerticalDiagonalText(displayValues.Count, interval, width, margin, options.ChartSize.Height))
                {
                    return new AxisScale()
                    {
                        MajorInterval = interval,
                        MinorInterval = 1,
                        Min = 1,
                        Max = displayValues.Count,
                        TextOrientation = orientationCustom,
                        DisplayValues = displayValues
                    };
                }
                //}
            }
            else
            {
                var width = fontSize * Math.Sqrt(2);
                var margin = fontSize * 0.5;
                if (options.IsMulti || FitAsVerticalDiagonalText(displayValues.Count, interval, width, margin, plotAreaWidth)) //Check diagonal
                {
                    if (FitAsHorizontalText(displayValues, interval, shaper, fontSize, plotAreaWidth, options.IsMulti)) //Check horizontal
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
            List<object> displayValues = GetUniqueValues(values, options.IsMulti).Select(x => (object)x.ToString()).ToList();
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
                TextOrientation = (options.Axis.TextBody?.RotationAdjustedForTextVertical??0D) ==0 ? eTextOrientation.Horizontal : eTextOrientation.Custom,
                DisplayValues = displayValues
            };
        }


        private static bool FitAsHorizontalText(List<object> displayValues, int interval, ITextShaper shaper, float fontSize, double plotAreaWidth, bool isMultLevel)
        {
            var margin = fontSize * 0.3;
            double maxWidth;
            var pos = interval;
            if (isMultLevel)
            {
                var arr = (object[])displayValues[0];
                var lastIx = arr.Length - 1;
                maxWidth = GetTextWidth(shaper, arr[lastIx].ToString(), fontSize) + margin;
                while (pos < displayValues.Count && (maxWidth * displayValues.Count / interval) < plotAreaWidth)
                {
                    arr = (object[])displayValues[pos];
                    var width = GetTextWidth(shaper, arr[lastIx].ToString(), fontSize) + margin;
                    if (width > maxWidth)
                    {
                        maxWidth = width;
                        if ((maxWidth * displayValues.Count / interval) > plotAreaWidth) return false;
                    }
                    pos += interval;
                }
            }
            else
            {
                maxWidth = GetTextWidth(shaper, displayValues[0].ToString(), fontSize) + margin;
                while (pos < displayValues.Count && (maxWidth * displayValues.Count / interval) < plotAreaWidth)
                {
                    var width = GetTextWidth(shaper, displayValues[pos].ToString(), fontSize) + margin;
                    if (width > maxWidth)
                    {
                        maxWidth = width;
                        if ((maxWidth * displayValues.Count / interval) > plotAreaWidth) return false;
                    }
                    pos += interval;
                }
            }
            return true;
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

        internal static List<object> GetUniqueValues(List<object> values, bool isMulti)
        {
            var ret = new List<object>();
            var hs = new HashSet<string>();
            string[] prevNonEmpty=null;
            foreach (var v in values)
            {
                if (v is object[] array)
                {
                    if(isMulti)                    
                    {
                        if(prevNonEmpty==null) prevNonEmpty = new string[array.Length - 1]; //For empty values, we will use previous non empty value in the key to match up hierarki values.
                        var key = "";
                        var o = new object[array.Length];
                        for (int ix = 0;ix<array.Length;ix++)                        
                        {
                            var s = (object[])array[ix];
                            if(ix < array.Length-1 && string.IsNullOrEmpty(s[0].ToString()))
                            {
                                key += prevNonEmpty[ix].ToString() + s[1].ToString();
                            }
                            else
                            {
                                key += s[0].ToString() + s[1].ToString();
                                if (ix < array.Length - 1) prevNonEmpty[ix] = s[0].ToString();
                            }
                            o[ix] = ((object[])array[ix])[0];
                        }
                        if (hs.Add(key))
                        {
                            ret.Add(o);
                        }
                    }
                    else
                    {
                        var key = array[0].ToString() + array[1].ToString();
                        if (hs.Add(key))
                        {
                            ret.Add(array[0]);
                        }
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