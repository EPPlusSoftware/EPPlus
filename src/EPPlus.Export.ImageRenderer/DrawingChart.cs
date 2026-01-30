/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/

using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlusImageRenderer
{
    internal abstract class DrawingChart : DrawingBase
    {       
        public DrawingChart(ExcelChart chart) : base(chart) 
        {
            Chart = chart;
        }
        public ExcelChart Chart { get; set; }

        protected List<object> GetAxisValue(ExcelChartAxisStandard ax, RenderItem rect, out double? min, out double? max, out double? majorUnit)
        {
            var values = ax.GetAxisValues(out bool isCount);
            if(ax.AxisType == eAxisType.Cat && 
                isCount == false)
            {
                min = 0;
                max = values.Length;
                majorUnit = 1;
                return values.ToList();
            }
            var l = new List<object>();
            min = double.MaxValue;
            max = double.MinValue;
            foreach (var v in values)
            {
                var d = ConvertUtil.GetValueDouble(v, false, true);
                if(double.IsNaN(d))
                {
                    d = 0;
                }
                if(min>d)
                {
                    min = d;
                }
                if(max<d)
                {
                    max = d;
                }
            }
            var maxMajorTickmarks = 10; //TODO: Calculate based on rect size
            GetAutoMinMaxValue(ax, maxMajorTickmarks, isCount, ref min, ref max, out majorUnit);
            for(var v=min; v<=max;v+=majorUnit)  
            {
                l.Add(v);
            }
            return l;
            }

        private void GetAutoMinMaxValue(ExcelChartAxisStandard ax, int maxMajorTickmarks, bool isCount, ref double? min, ref double? max, out double? majorUnit)
        {
            if(ax.MinValue.HasValue)
            {
                min = ax.MinValue;
            }
            else
            {
                if (isCount)
                {
                    min = 1;
                }
                else
                {
                    var diffFromZero = (max - min) / max;
                    if (diffFromZero > 0.091)
                    {
                        min = 0;
                    }
                }
            }

            if(isCount)
            {
                majorUnit = 1;
            }
            else
            {
                if (ax.MaxValue.HasValue)
                {
                    max = ax.MaxValue;
                    majorUnit = ax.MajorUnit ?? GetAutoUnit(min.Value, max.Value);
                    if (ax.MinValue.HasValue == false)
                    {
                        var newMin = max - majorUnit;
                        while (newMin > min)
                        {
                            newMin -= majorUnit.Value;
                        }
                        min = newMin;
                    }
                }
                else
                {
                    majorUnit = ax.MajorUnit ?? GetAutoUnit(min.Value, max.Value);
                    if (isCount == false)
                    {
                        var diff = max.Value - min.Value;
                        var newMax = min.Value + majorUnit;
                        while ((newMax - min) < (diff * 1.05))
                        {
                            newMax += majorUnit.Value;
                        }
                        max = newMax;
                    }
                    if(min != 0 && max-min<9)
                    {
                        min -= 2;
                    }
                }
                var newUnit = majorUnit;
                while (newUnit >= 2 && (max - min) / newUnit > maxMajorTickmarks)
                {
                    newUnit /= 2;
                }
            }
        }

        private double GetAutoUnit(double min, double max)
        {
            var diff = max - min;
            if (diff < 8)
            {
                return 1;
            }
            else
            {
                var rawMajorUnit = diff;
                var exponent = Math.Floor(Math.Log10(rawMajorUnit));
                var fraction = rawMajorUnit / (Math.Pow(10, exponent));
                double unit;
                if (fraction <= 1)
                {
                    unit = 1D;
                }
                else if (fraction <= 2)
                {
                    unit = 2;
                }
                else if (fraction <= 2.5)
                {
                    unit = 2.5;
                }
                else if (fraction <= 5)
                {
                    unit = 5;
                }
                else
                {
                    unit = 10;
                }

                var axMax = unit * Math.Pow(10, exponent);
                var axMin = Math.Floor(min / axMax) * axMax;
                axMax = Math.Ceiling(max / axMax) * axMax;
                return axMax / 10;
            }
        }
    }
}