using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Utils.DateUtils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.ImageRenderer.Svg.Chart.Util
{
    internal class DateAxisScaleCalculator
    {


        /// <summary>
        /// Calculate automatic axis settings for date data
        /// </summary>
        internal static AxisScale Calculate(double dataMin, double dataMax, double chartHeightPixels, AxisOptions axisOptions = null)
        {

            ComputeAutoAxis(DateTime.FromOADate(dataMin), DateTime.FromOADate(dataMax));
            var range = dataMax - dataMin;

            // Calculate major majorUnit based on range
            CalculateMajorUnit(axisOptions.LockedMin ?? dataMin, axisOptions.LockedMax ?? dataMax, out double majorValue, out double axisMin, out double axisMax, out double minorUnit, out eTimeUnit majorUnit);
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
                MinorDateUnit = majorUnit
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
        public static AxisResult ComputeAutoAxis(DateTime dataMin, DateTime dataMax)
        {
            double serialMin = ToExcelSerial(dataMin);
            double serialMax = ToExcelSerial(dataMax);

            // Step 1: pick a nice major unit
            double rawInterval = (serialMax - serialMin) / 10.0;
            int interval = NiceIntervals.FirstOrDefault(v => v >= rawInterval);
            if (interval == 0)
                interval = 730; // fallback for very large ranges

            // Step 2: snap outward to multiples of the interval
            double axisMin = Math.Floor(serialMin / interval) * interval;
            double axisMax = Math.Ceiling(serialMax / interval) * interval;

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
        private static void CalculateMajorUnit(double min, double max,out double majorValue, out double axisMin, out double axisMax, out double minorUnits, out eTimeUnit majorUnit)
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


            axisMin = Math.Floor(min / interval) * interval;
            axisMax = Math.Ceiling(max / interval) * interval;
            if(axisMax==max)
            {
                axisMax+= interval;
            }

            // Step 3: expand until we have enough ticks, then re-snap
            while ((axisMax - axisMin) / interval < TargetTicks - 1 && ((min-axisMin) < interval * 2))
            {
                axisMin -= majorValue;
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
    }
}