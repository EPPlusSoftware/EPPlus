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
            var range = dataMax - dataMin;

            // Calculate major majorUnit based on range
            CalculateMajorUnit(dataMin, dataMax, out double majorValue, out double minorUnit, out eTimeUnit majorUnit);

            return new AxisScale
            {
                Min = dataMin,
                Max = dataMax,
                MajorInterval = majorValue,
                MajorDateUnit = majorUnit,
                MinorDateUnit = majorUnit
            };
        }

        /// <summary>
        /// Calculate major majorUnit based on data range (in days)
        /// </summary>
        private static void CalculateMajorUnit(double min, double max,out double majorValue, out double minorUnits, out eTimeUnit majorUnit)
        {                        
            var dtMin = DateTime.FromOADate(ExcelNormalizeOADate(min));
            var dtMax = DateTime.FromOADate(ExcelNormalizeOADate(max));

            var rangeDays = (dtMax - dtMin).TotalDays;
            // Target approximately 6 intervals
            const int targetIntervals = 6;
            var roughUnit = rangeDays / targetIntervals;

            // Return nice round numbers based on the rough majorUnit
            if (rangeDays <= 1)
            {
                majorValue = 1;
                majorUnit = eTimeUnit.Days;
            }
            else if (roughUnit < 7)
            {
                // Multiple days - round to nearest day
                majorValue = Math.Max(1, Math.Round(roughUnit));
                majorUnit = eTimeUnit.Days;
            }
            else if (roughUnit < 14)
            {
                majorValue = 7;
                majorUnit = eTimeUnit.Days;
            }
            else if (roughUnit < 21)
            {
                majorValue = 14;
                majorUnit = eTimeUnit.Days;
            }
            else if (roughUnit < 60)
            {
                majorValue = 1;
                majorUnit = eTimeUnit.Months;
            }
            else if (roughUnit < 120)
            {
                majorValue = 2;
                majorUnit = eTimeUnit.Months;
            }
            else if (roughUnit < 180)
            {
                majorValue = 3;
                majorUnit = eTimeUnit.Months;
            }
            else if (roughUnit < 365)
            {
                majorValue = 6;
                majorUnit = eTimeUnit.Months;
            }
            else if (roughUnit < 730)
            {
                majorValue = 1;
                majorUnit = eTimeUnit.Years;
            }
            else if (roughUnit < 1825)
            {
                // Multiple years
                majorValue = Math.Ceiling(roughUnit / 365);
                majorUnit = eTimeUnit.Years;
            }
            else
            {
                // 5-year increments for very long ranges
                majorValue = Math.Ceiling(roughUnit / 365 / 5) * 5;
                majorUnit = eTimeUnit.Years;
            }
            minorUnits = 1;
        }

        private static double ExcelNormalizeOADate(double value)
        {
            return Math.Min(Math.Max(value, 0), 2958465.99999999);
        }
    }
}