using EPPlus.Export.Renderer;
using OfficeOpenXml;
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
    internal class DateAxisScaleCalculator
    {
        /// <summary>
        /// Calculate automatic axis settings for date data
        /// </summary>
        internal static AxisScale Calculate(double dataMin, double dataMax, AxisOptions axisOptions = null)
        {

            ComputeAutoAxis(DateTime.FromOADate(dataMin), DateTime.FromOADate(dataMax), axisOptions.AddPadding);
            var range = dataMax - dataMin;

            // Calculate major majorUnit based on range
            CalculateMajorUnit(axisOptions.LockedMin ?? dataMin, axisOptions.LockedMax ?? dataMax, axisOptions.AddPadding, out double majorValue, out double axisMin, out double axisMax, out double minorUnit, out eTimeUnit majorUnit);
            if(axisOptions.LockedInterval.HasValue)
            {
                majorValue = axisOptions.LockedInterval.Value;
            }

            return new AxisScale
            {
                Min = axisMin,
                Max = axisMax,
                MajorInterval = majorValue,
                MajorDateUnit = majorUnit,
                MinorDateUnit = majorUnit,
                TextOrientation = eTextOrientation.Horizontal
            };
        }
        // Excel's date serial epoch: Dec 30, 1899
        private static readonly DateTime ExcelEpoch = new DateTime(1899, 12, 30);

        // Excel legacy leap year bug: dates after Feb 28, 1900 are offset by 1
        private static readonly DateTime ExcelLeapBugThreshold = new DateTime(1900, 3, 1);

        // Excel's "nice" interval ladder in days
        private static readonly int[] NiceIntervals = { 1, 2, 5, 7, 14, 30, 91, 182, 365, 730 };
        private const int TargetTicks = 11; // Excel typically targets 10–12 ticks
        public class AxisResult
        {
            public DateTime AxisMin { get; set; }
            public DateTime AxisMax { get; set; }
            public int IntervalDays { get; set; }
            public double AxisMinSerial { get; set; }
            public double AxisMaxSerial { get; set; }
            public int TickCount => (int)Math.Round((AxisMaxSerial - AxisMinSerial) / IntervalDays) + 1;

            public override string ToString() =>
                $"Axis Min : {AxisMin:yyyy-MM-dd} (serial {AxisMinSerial})\n" +
                $"Axis Max : {AxisMax:yyyy-MM-dd} (serial {AxisMaxSerial})\n" +
                $"Interval : {IntervalDays} day(s)\n" +
                $"Ticks    : {TickCount}";
        }
        /// <summary>
        /// Converts a DateTime to an Excel serial number (days since Dec 30, 1899),
        /// accounting for Excel's 1900 leap year bug.
        /// </summary>
        public static double ToExcelSerial(DateTime date)
        {
            double serial = (date - ExcelEpoch).TotalDays;
            if (date >= ExcelLeapBugThreshold)
                serial += 1; // Excel's legacy off-by-one
            return serial;
        }

        /// <summary>
        /// Converts an Excel serial number back to a DateTime.
        /// </summary>
        public static DateTime FromExcelSerial(double serial)
        {
            if (serial >= 61) // 61 = the phantom Feb 29, 1900 in Excel
                serial -= 1;
            return ExcelEpoch.AddDays(serial);
        }
        /// <summary>
        /// Computes Excel-compatible auto axis min, max, and major unit for a date axis.
        /// If the OOXML already contains explicit min/max values, use those directly instead.
        /// </summary>
        public static AxisResult ComputeAutoAxis(DateTime dataMin, DateTime dataMax, bool snapInterval)
        {
            double serialMin = ToExcelSerial(dataMin);
            double serialMax = ToExcelSerial(dataMax);

            // Step 1: pick a nice major unit
            double rawInterval = (serialMax - serialMin) / 10.0;
            int interval = NiceIntervals.FirstOrDefault(v => v >= rawInterval);
            if (interval == 0)
                interval = 730; // fallback for very large ranges

            double axisMin = serialMin;
            double axisMax = serialMax;
            if (snapInterval)
            {
                // Step 2: snap outward to multiples of the interval
                axisMin = Math.Floor(serialMin / interval) * interval;
                axisMax = Math.Ceiling(serialMax / interval) * interval;
            }
            else
            {
                axisMin = serialMin;
                axisMax = serialMax;
            }
            // Step 3: expand until we have enough ticks, then re-snap
            while ((axisMax - axisMin) / interval < TargetTicks - 1)
            {
                axisMin -= interval;
                axisMax += interval;
            }

            axisMin = Math.Floor(axisMin / interval) * interval;
            axisMax = Math.Ceiling(axisMax / interval) * interval;

            return new AxisResult
            {
                AxisMin = FromExcelSerial(axisMin),
                AxisMax = FromExcelSerial(axisMax),
                IntervalDays = interval,
                AxisMinSerial = axisMin,
                AxisMaxSerial = axisMax,
            };
        }
        /// <summary>
        /// Calculate major majorUnit based on data range (in days)
        /// </summary>
        private static void CalculateMajorUnit(double min, double max, bool snapToInterval, out double majorValue, out double axisMin, out double axisMax, out double minorUnits, out eTimeUnit majorUnit)
        {                        
            var dtMin = DateTime.FromOADate(ExcelNormalizeOADate(min));
            var dtMax = DateTime.FromOADate(ExcelNormalizeOADate(max));

            var rangeDays = (dtMax - dtMin).TotalDays;
            // Target approximately 10 intervals
            const int targetIntervals = 9;
            var roughUnit = rangeDays / targetIntervals;
            double interval;
            // Return nice round numbers based on the rough majorUnit
            if (roughUnit < 1)
            {
                interval = majorValue = 1;
                majorUnit = eTimeUnit.Days;
            }
            else if (roughUnit < 2)
            {
                interval = majorValue = 2;
                majorUnit = eTimeUnit.Days;
            }
            else if (roughUnit < 5)
            {
                // Multiple days - round to nearest day
                interval = majorValue = 5;
                majorUnit = eTimeUnit.Days;
            }
            else if (roughUnit < 7)
            {
                // Multiple days - round to nearest day
                interval = majorValue = 7;
                majorUnit = eTimeUnit.Days;
            }
            else if (roughUnit < 14)
            {
                interval = majorValue = 14;
                majorUnit = eTimeUnit.Days;
            }
            else if (roughUnit < 21)
            {
                interval = majorValue = 21;
                majorUnit = eTimeUnit.Days;
            }
            else if (roughUnit < 60)
            {
                majorValue = 1;
                interval = 30;
                majorUnit = eTimeUnit.Months;
            }
            else if (roughUnit < 120)
            {
                majorValue = 2;
                interval = 60;
                majorUnit = eTimeUnit.Months;
            }
            else if (roughUnit < 180)
            {
                majorValue = 3;
                interval = 90;
                majorUnit = eTimeUnit.Months;
            }
            else if (roughUnit < 365)
            {
                majorValue = 6;
                interval = 180;
                majorUnit = eTimeUnit.Months;
            }
            else if (roughUnit < 730)
            {
                majorValue = 1;
                interval = 365;
                majorUnit = eTimeUnit.Years;
            }
            else if (roughUnit < 1825)
            {
                // Multiple years
                majorValue = Math.Ceiling(roughUnit / 365);
                interval = 365 * majorValue;
                majorUnit = eTimeUnit.Years;
            }
            else
            {
                // 5-year increments for very long ranges
                majorValue = Math.Ceiling(roughUnit / 365 / 5) * 5;
                interval = 365 * majorValue;
                majorUnit = eTimeUnit.Years;
            }

            if (snapToInterval)
            {
                axisMin = Math.Floor(min / interval) * interval;
                axisMax = Math.Ceiling(max / interval) * interval;

                if (axisMax == max)
                {
                    axisMax += interval;
                }
                // Step 3: expand until we have enough ticks, then re-snap
                while ((axisMax - axisMin) / interval < TargetTicks - 1 && ((min - axisMin) < interval * 2))
                {
                    axisMin -= majorValue;
                }
            }
            else
            {
                axisMin= min;
                axisMax = max;
            }

             minorUnits = 1;
        }

        private static double ExcelNormalizeOADate(double value)
        {
            return Math.Min(Math.Max(value, 0), 2958465.99999999);
        }

        private static DateTime FloorToUnit(double value, eTimeUnit unit, int interval)
        {
            var date = DateTime.FromOADate(value);
            switch (unit)
            {
                case eTimeUnit.Days:
                    if(interval % 7 != 0)
                    {
                        return date.Date.AddDays(-(((int)date.DayOfWeek + 6) % 7));
                    }
                    else
                    {
                        return date;
                    }
                        //return date.Date.AddDays(-(date.DayOfYear % interval));
                case eTimeUnit.Months:
                    return new DateTime(date.Year, date.Month, 1)
                        .AddMonths(-((date.Month - 1) % interval));

                case eTimeUnit.Years:
                    return new DateTime(date.Year - (date.Year % interval), 1, 1);
                default:
                    return date;
            };
        }

        private static DateTime CeilingToUnit(double value, eTimeUnit unit, int interval)
        {
            DateTime floor = FloorToUnit(value, unit, interval);
            
            var date = DateTime.FromOADate(value);
            if (floor == date) return date; // already on boundary

            switch (unit)
            {
                case eTimeUnit.Days:
                    return floor.AddDays(interval);
                case eTimeUnit.Months:
                    return floor.AddMonths(interval);
                case eTimeUnit.Years:
                    return floor.AddYears(interval);
                default:
                    return date;
            };
        }

        internal static AxisScale CalculateByWidthAllowDiagonal(List<object> values, double min, double max, ITextMeasurer tm, AxisOptions options)
        {
            var ax = options.Axis;
            var plotAreaWidth = options.ChartSize.Bounds.Width;
            var mf = ax.Font.GetMeasureFont();
            int interval;
            eTimeUnit unit;
            string format = GetNumberFormat(options);
            var minString = DateTime.FromOADate(min).ToString(format);
            var res = tm.MeasureText(minString, mf);
            if (options.LockedInterval.HasValue)
            {
                interval = (int)options.LockedInterval.Value;
                unit = options.LockedIntervalUnit ?? eTimeUnit.Days;
            }
            else
            {
                GetStartInterval(values.Count, max - min, out interval, out unit);
                if(ax.TextBody.Rotation.HasValue==false || ax.TextBody.Rotation==-1000)
                {
                    //Get interval for maximum width with vertical text.
                    while (FitAsVerticalDiagonalText(min, max, interval, unit, res.Height, res.Height * 0.3, plotAreaWidth) == false)
                    {
                        AddIntervall(ref interval, ref unit);
                    }

                    //Get max text width when using diagonal text
                    var width = mf.Size * Math.Sqrt(2);
                    var margin = mf.Size * 0.5;

                    if (FitAsVerticalDiagonalText(min, max, interval, unit, width, margin, plotAreaWidth)) //Check diagonal
                    {
                        if (FitAsHorizontalText(tm, options, min, max, interval, unit, res.Height, plotAreaWidth)) //Check horizontal
                        {
                            return new AxisScale()
                            {
                                MajorInterval = interval,
                                MinorInterval = 1,
                                MinorDateUnit = unit,
                                MajorDateUnit = unit,
                                Min = min,
                                Max = max,
                                TextOrientation = eTextOrientation.Horizontal,
                            };
                        }
                        else
                        {
                            return new AxisScale()
                            {
                                MajorInterval = interval,
                                MinorInterval = 1,
                                MinorDateUnit = unit,
                                MajorDateUnit = unit,
                                Min = min,
                                Max = max,
                                TextOrientation = eTextOrientation.Diagonal,
                            };
                        }
                    }
                }
                else
                {
                    var rot = ax.TextBody.Rotation.Value % 360;
                    var sin = Math.Sin(MathHelper.Radians(rot));
                    var cos = Math.Cos(MathHelper.Radians(rot));
                    var width = res.Width * cos + res.Height * sin;
                    //Get interval for maximum width with vertical text.
                    while (FitAsVerticalDiagonalText(min, max, interval, unit, width, width * 0.3, plotAreaWidth) == false)
                    {
                        AddIntervall(ref interval, ref unit);
                    }
                    eTextOrientation orientation;
                    switch(rot)
                    {
                        case 45:
                        case 315:
                            orientation = eTextOrientation.Diagonal;
                            break;
                        case 90:
                        case 270:
                            orientation = eTextOrientation.Vertical;
                            break;
                        case 0:
                        case 180:
                            orientation = eTextOrientation.Horizontal;
                            break;
                        default:
                            orientation = eTextOrientation.Custom;
                            break;
                    }
                    return new AxisScale()
                    {
                        MajorInterval = interval,
                        MinorInterval = 1,
                        MinorDateUnit = unit,
                        MajorDateUnit = unit,
                        Min = min,
                        Max = max,
                        TextOrientation = orientation,
                    };
                }
            }

            return new AxisScale()
            {
                MajorInterval = interval,
                MinorInterval = 1,
                MinorDateUnit = unit,
                MajorDateUnit = unit,
                Min = min,
                Max = max,
                TextOrientation = ax.TextBody.VerticalText == OfficeOpenXml.Drawing.eTextVerticalType.Horizontal ? eTextOrientation.Horizontal : eTextOrientation.Vertical,
            };
        }

        private static string GetNumberFormat(AxisOptions options)
        {
            var format = options.NumberFormat;
            if (string.IsNullOrEmpty(format))
            {
                format = options.Axis.FormatOrFirstValueFormat;
                if (string.IsNullOrEmpty(format))
                {
                    format = "yyyy-MM-dd";
                }
            }

            return format;
        }

        private static void GetStartInterval(double count, double span, out int interval, out eTimeUnit unit)
        {
            var blockSize = span / count;
            if(blockSize<3.5)
            {
                interval = 1;
                unit = eTimeUnit.Days;
            }
            if(blockSize < 15)
            {
                interval = 7;
                unit = eTimeUnit.Days;
            }
            else if(blockSize < 180)
            {
                interval = 1;
                unit = eTimeUnit.Months;
            }
            else
            {
                interval = 1;
                unit = eTimeUnit.Years;
            }
        }

        //private static bool FitAsDiagonalText(double min, double max, int interval, eTimeUnit unit, float size, double plotAreaWidth)
        //{
        //    var margin = size * 0.5;
        //    var width = size * Math.Sqrt(2) + margin;

        //    if (unit == eTimeUnit.Days)
        //    {
        //        return ((max - min) / interval) * width < plotAreaWidth;
        //    }
        //    else if(unit == eTimeUnit.Months)
        //    {
        //        w
        //    }
        //}

        private static void AddIntervall(ref int interval, ref eTimeUnit unit)
        {
            if (unit == eTimeUnit.Days)
            {
                switch (interval)
                {
                    case 1:
                        interval = 2;
                        break;
                    case 2:
                        interval = 7; //Week
                        break;
                    case 7:
                        interval = 1; //Month
                        unit = eTimeUnit.Months;
                        break;
                    default:
                        interval *= 10;
                        break;
                }
            }
            else if (unit == eTimeUnit.Months)
            {
                unit = eTimeUnit.Years;
            }
            else if (unit == eTimeUnit.Years)
            {
                interval++;
            }
        }

        private static bool FitAsHorizontalText(ITextMeasurer tm, AxisOptions options, double min, double max, int interval, eTimeUnit unit, float height, double width)
        {
            var minMargin = 2; //2 Points
            var minDate = DateTime.FromOADate(min);
            var maxDate = DateTime.FromOADate(max);
            var date = minDate;
            var horizontalWidth = 0D;
            var nf = options.NumberFormat;
            var mf = options.Axis.Font.GetMeasureFont();
            while (date < maxDate)
            {
                var textWidth = tm.MeasureText(date.ToString(), mf).Width;
                horizontalWidth += textWidth + minMargin;
                if(horizontalWidth > width)
                {
                    return false;
                }
                if (unit == eTimeUnit.Days)
                {
                    date = date.AddDays(interval);
                }
                else if(unit == eTimeUnit.Months)
                {
                    date = minDate.AddDays(1).AddMonths(interval).AddDays(-1); //Extra adds to keep last day of month
                }
                else
                {
                    date = minDate.AddDays(1);
                }
            }
            return true;
        }

        private static bool FitAsVerticalDiagonalText(double min, double max, int interval,eTimeUnit unit, double textWidth, double margin, double plotAreaWidth)
        {
            if (unit == eTimeUnit.Days)
            {
                var items = ((max - min) / interval);
                return  items * textWidth + (items-1) * margin < plotAreaWidth;
            }
            else if(unit== eTimeUnit.Months) 
            {
                var minDate = DateTime.FromOADate(min);
                var maxDate = DateTime.FromOADate(max);
                var date = minDate.AddDays(1).AddMonths(interval).AddDays(-1); //Handles last day of month.
                var items = 1;
                while(date <= maxDate)
                {
                    items++;
                    if (items * textWidth + (items - 1) * margin > plotAreaWidth) return false;
                    date = date.AddDays(1).AddMonths(interval).AddDays(-1);
                }
                return items * textWidth + (items - 1) * margin < plotAreaWidth;
            }
            else
            {
                var minDate = DateTime.FromOADate(min);
                var maxDate = DateTime.FromOADate(max);
                var date = minDate.AddDays(1).AddYears(interval).AddDays(-1); //Handles leap year
                var items = 1;
                while (date <= maxDate)
                {
                    items++;
                    if (items * textWidth + (items - 1) * margin > plotAreaWidth) return false;
                    date = date.AddDays(1).AddYears(interval).AddDays(-1);
                }
                return items * textWidth + (items - 1) * margin < plotAreaWidth;

            }
        }

        internal static AxisScale CalculateByWidthHeight(double widthOrHeight, double min, double max, ITextMeasurer tm, AxisOptions options)
        {
            var ax = options.Axis;
            var mf = ax.Font.GetMeasureFont();
            int interval;
            eTimeUnit unit;
            var minString = DateTime.FromOADate(min).ToString(options.NumberFormat);
            var res = tm.MeasureText(minString, mf);
            if (options.LockedInterval.HasValue)
            {
                interval = (int)options.LockedInterval.Value;
                unit = options.LockedIntervalUnit ?? eTimeUnit.Days;
            }
            else
            {
                interval = 1;
                unit = eTimeUnit.Days;
                var range = max - min;

                double axis_min, axis_max;
                if (options.AddPadding)
                {
                    axis_min = options.LockedMin ?? Math.Floor((min - 0.5 * range) / interval) * interval;
                    axis_max = options.LockedMax ?? Math.Ceiling((max + 0.2 * range) / interval) * interval;
                }
                else
                {
                    axis_min = min;
                    axis_max = max;
                }
                //Get interval for maximum width with vertical text.
                while (FitAsVerticalDiagonalText(axis_min, axis_max, interval, unit, res.Height, res.Height * 0.3, widthOrHeight) == false)
                {
                    AddIntervall(ref interval, ref unit);
                    double days;
                    switch (unit)
                    {
                        case eTimeUnit.Months:
                            days = 30 * interval;
                            break;
                        case eTimeUnit.Years:
                            days = 365 * interval;
                            break;
                        default:
                            days = interval;
                            break;
                    }

                    if (options.AddPadding)
                    {
                        axis_min = options.LockedMin ?? Math.Floor((min - 0.5 * range) / days) * days;
                        axis_max = options.LockedMax ?? Math.Ceiling((max + 0.2 * range) / days) * days;
                    }
                }
                min = axis_min;
                max = axis_max;
            }

            return new AxisScale()
            {
                MajorInterval = interval,
                MinorInterval = 1,
                MinorDateUnit = unit,
                MajorDateUnit = unit,
                Min = min,
                Max = max,
                TextOrientation = eTextOrientation.Horizontal
            };

        }
    }
}