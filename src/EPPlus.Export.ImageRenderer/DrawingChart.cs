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
                max = values.Length-1;
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
            GetAutoMinMaxValue(ax, isCount, ref min, ref max, out majorUnit);
            for(var v=min; v<=max;v+=majorUnit)  
            {
                l.Add(v);
            }
            return l;
        }

        private void GetAutoMinMaxValue(ExcelChartAxisStandard ax, bool isCount, ref double? min, ref double? max, out double? majorUnit)
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
            if(ax.MaxValue.HasValue)
            {
                max = ax.MaxValue;
                majorUnit = ax.MajorUnit ?? GetAutoUnit(min.Value, max.Value);
            }
            else
            {
                majorUnit = ax.MajorUnit ?? GetAutoUnit(min.Value, max.Value);
                if (isCount==false)
                {
                    var diff = max.Value - min.Value;
                    var newMax = max.Value + majorUnit;
                    while ((newMax - min) < (diff * 1.05))
                    {
                        newMax += majorUnit.Value;
                    }
                    max = newMax;
                }
            }
        }

        private double GetAutoUnit(double min, double max)
        {
            var diff = max - min;
            var pow = 0;
            while (diff < 10)
            {
                diff *= 10;
                pow++;
            }
            var unit = diff / 10D;
            var rest = unit % 10d;
            if (rest < 5)
            {
                unit -= rest;
            }
            else
            {
                unit += 10 - rest;
            }
            if(pow>0)
            {
                unit /= Math.Pow(10, pow);
            }
            if (unit == 0)
            {
                return 1;
            }
            return unit;
        }
    }
}