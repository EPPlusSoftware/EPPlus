using System;
using System.Collections.Generic;

internal class ValueAxisScaleCalculator
{
    internal class AxisOptions
    {
        public double? LockedMin { get; set; }
        public double? LockedMax { get; set; }
        public double? LockedInterval { get; set; }
        public bool AddPadding { get; set; } = false;
    }
    internal class AxisScale
    {
        public double Min { get; set; }
        public double Max { get; set; }
        public double Interval { get; set; }
        public int TickCount { get; set; }
    }

    internal static AxisScale Calculate(double dataMin, double dataMax, double chartHeightPixels, AxisOptions axisOptions = null)
    {
        int desiredTicks = chartHeightPixels < 200 ? 4 :
                             chartHeightPixels < 400 ? 5 : 6;

        if (axisOptions == null)
        { 
            axisOptions = new AxisOptions();
        }

        // Handle equal data series.
        if (dataMin == dataMax)
        {
            if (dataMin < 0)
            {
                dataMax = 0;
            }
            else
            {
                dataMin = 0;
            }
        }

        var isAllPositive = dataMin > 0 && dataMax > 0;
        var isAllNegativ = dataMin < 0 && dataMax < 0;

        double interval;
        if (axisOptions.LockedInterval.HasValue)
        {
            interval = axisOptions.LockedInterval.Value;

            // Calculate min and max based on locked interval
            if (!axisOptions.LockedMin.HasValue)
            {
                if (axisOptions.AddPadding)
                {
                    dataMin = Math.Floor(dataMin * 0.1 / interval) * interval;
                    if(isAllPositive && dataMin < 0 )
                    {
                        dataMin = 0;
                    }
                }
            }
            else
            {
                dataMin = axisOptions.LockedMin.Value;
            }

            if (!axisOptions.LockedMax.HasValue)
            {
                dataMax = Math.Ceiling(dataMax * 0.1 / interval) * interval;
            } 
            else
            {
                if (axisOptions.AddPadding)
                {
                    dataMax = axisOptions.LockedMax.Value;
                    if (isAllNegativ && dataMax > 0)
                    {
                        dataMax = 0;
                    }
                }
            }

            int tickCount = (int)Math.Round((dataMax - dataMin) / Math.Min(desiredTicks, interval)) + 1;
            return new AxisScale    
            {
                Min = dataMin,
                Max = dataMax,
                Interval = interval,
                TickCount = tickCount
            };
        }
        else
        {
            if (axisOptions.LockedMin.HasValue)
            {
                dataMin = axisOptions.LockedMin.Value;
            }
            if (axisOptions.LockedMax.HasValue)
            {
                dataMax = axisOptions.LockedMax.Value;
            }

            double dataRange = dataMax - dataMin;

            // Add padding (10%)
            double paddedMin = dataMin - (dataRange * 0.1);
            double paddedMax = dataMax + (dataRange * 0.1);

            //Normalize to zero if all data is positive or negative
            if (paddedMin < 0 && isAllPositive)
            {
                paddedMin = 0;
            }
            if (paddedMax > 0 && isAllNegativ)
            {
                paddedMax = 0;
            }

            // Calculate nice interval
            double roughInterval = (paddedMax - paddedMin) / desiredTicks;
            double scaleInterval = GetScaleNumber(roughInterval, true);

            // Calculate false min and max
            double axisMin = Math.Floor(paddedMin / scaleInterval) * scaleInterval;
            double axisMax = Math.Ceiling(paddedMax / scaleInterval) * scaleInterval;

            // Calculate actual tick count
            int tickCount = (int)Math.Round((axisMax - axisMin) / scaleInterval) + 1;
            return new AxisScale
            {
                Min = axisOptions.LockedMin ?? axisMin,
                Max = axisOptions.LockedMax ?? axisMax,
                Interval = scaleInterval,
                TickCount = tickCount
            };
        }
    }

    private static double GetScaleNumber(double value, bool round)
    {
        double exponent = Math.Floor(Math.Log10(value));
        double fraction = value / Math.Pow(10, exponent);
        double scaleFraction;

        if (round)
        {
            if (fraction < 1.5) scaleFraction = 1;
            else if (fraction < 3) scaleFraction = 2;
            else if (fraction < 7) scaleFraction = 5;
            else scaleFraction = 10;
        }
        else
        {
            if (fraction <= 1) scaleFraction = 1;
            else if (fraction <= 2) scaleFraction = 2;
            else if (fraction <= 5) scaleFraction = 5;
            else scaleFraction = 10;
        }

        return scaleFraction * Math.Pow(10, exponent);
    }

    public static List<double> GetTickValues(AxisScale scale)
    {
        var ticks = new List<double>();
        for (int i = 0; i < scale.TickCount; i++)
        {
            double tickValue = scale.Min + (i * scale.Interval);
            ticks.Add(Math.Round(tickValue, 10)); // Round to avoid floating point errors
        }
        return ticks;
    }
}