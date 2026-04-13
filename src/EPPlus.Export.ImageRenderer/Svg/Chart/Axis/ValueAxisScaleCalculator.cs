using Microsoft.VisualBasic;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;

internal class ValueAxisScaleCalculator
{
    internal static AxisScale Calculate(double dataMin, double dataMax, double chartHeightPixels, AxisOptions axisOptions = null)
    {
        return GetNumberScale(ref dataMin, ref dataMax, chartHeightPixels, ref axisOptions);
    }

    private static AxisScale GetNumberScale(ref double dataMin, ref double dataMax, double chartHeightPixels, ref AxisOptions axisOptions)
    {
        var desiredTicks = chartHeightPixels < 200 ? 4 :
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

        var isAllPositive = dataMin >= 0 && dataMax >= 0;
        var isAllNegative = dataMin <= 0 && dataMax <= 0;
        if(dataMin < 0 && dataMax > 0 && axisOptions.IsStacked100)
        {
            desiredTicks *= 2;
        }
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
                    if(axisOptions.IsStacked100 && dataMin < -1)
                    {
                        dataMin = -1;
                    }
                    if (isAllPositive && dataMin < 0)
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
                if (axisOptions.IsStacked100 && dataMin > 1)
                {
                    dataMax = 1;
                }
            }
            else
            {
                if (axisOptions.AddPadding)
                {
                    dataMax = axisOptions.LockedMax.Value;
                    if (isAllNegative && dataMax > 0)
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
                MajorInterval = interval,
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

            if (axisOptions.LockedMin.HasValue==false && dataMin / dataMax > 0.666666)
            {
                dataMin = 0;
            }

            double dataRange = dataMax - dataMin;

            double axisMin;
            double axisMax;
            double roughInterval;
            double scaleInterval;

            if (axisOptions.AddPadding)
            {
                // Add padding (10%)
                if (axisOptions.LockedMin.HasValue)
                {
                    axisMin = axisOptions.LockedMin.Value;
                }
                else
                {
                    axisMin = dataMin - (dataRange * 0.05);
                    if (axisOptions.IsStacked100)
                    {
                        if (axisMin < -1)
                        {
                            axisMin = -1;
                        }
                        else if(axisMin>0)
                        {
                            axisMin = 0;
                        }
                    }
                    //Normalize to zero if all data is positive or negative
                    if (axisMin < 0 && isAllPositive)
                    {
                        axisMin = 0;
                    }
                }

                if (axisOptions.LockedMax.HasValue)
                {
                    axisMax = axisOptions.LockedMax.Value;
                }
                else
                {
                    axisMax = dataMax + (dataRange * 0.05);
                    if (axisOptions.IsStacked100 && axisMax > 1)
                    {
                        axisMax = 1;
                    }

                    if (axisMax > 0 && isAllNegative)
                    {
                        axisMax = 0;
                    }
                }

                // Calculate interval
                roughInterval = (axisMax - axisMin) / desiredTicks;
                scaleInterval = GetScaleNumber(roughInterval, true);

                if (!axisOptions.LockedMin.HasValue) axisMin = Math.Floor(axisMin / scaleInterval) * scaleInterval;
                if (!axisOptions.LockedMax.HasValue) axisMax = Math.Ceiling(axisMax / scaleInterval) * scaleInterval;
            }
            else
            {
                roughInterval = (dataMax - dataMin) / desiredTicks;
                scaleInterval = GetScaleNumber(roughInterval, true);
                if (dataMin / dataMax >= 0.666666666)
                {
                    axisMin = 0;
                }
                else
                {
                    axisMin = Math.Floor(dataMin / scaleInterval) * scaleInterval;
                }

                // Calculate interval

                axisMax = dataMax;
            }

            var tickCount = (int)Math.Round((axisMax - axisMin) / scaleInterval) + 1;
            return new AxisScale
            {
                Min = axisOptions.LockedMin ?? axisMin,
                Max = axisOptions.LockedMax ?? axisMax,
                MajorInterval = scaleInterval,
                MinorInterval = scaleInterval / 5,
                TickCount = tickCount,
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
            double tickValue = scale.Min + (i * scale.MajorInterval);
            ticks.Add(Math.Round(tickValue, 10)); // Round to avoid floating point errors
        }
        return ticks;
    }
}